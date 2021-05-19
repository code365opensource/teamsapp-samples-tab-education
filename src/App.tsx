import { BrowserRouter as Router, Route } from "react-router-dom";
import {
  Flex,
  Provider,
  teamsTheme,
  Text, Segment,
  Form,
  FormInput,
  FormButton,
  FormDropdown,
  List,
  FormTextArea,
} from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import React, { useEffect, useState } from "react";


function Parent() {
  const [parentInfo, setParentInfo] = useState<[any]>();//家长信息



  useEffect(() => {
    //默认尝试去加载localstorage里面的数据，并且从服务器刷新一次数据
    const email = localStorage.getItem("parentEmail");
    if (email) {
      fetch(`api/service?call=getParentInfoByEmail&email=${email}`).then(x => x.json()).then(x => {
        setParentInfo(x);
      });
    }
  }, []);


  // 作业列表组件
  const AssignmentList = (props: { class: string, student: string }) => {
    const [assignments, setAssignments] = useState<[any]>();//作业列表


    const AssignmentStatus = (props: { id: string, class: string, student: string }) => {
      const [status, setStatus] = useState<string>("正在检索状态");

      useEffect(() => {
        fetch(`api/service?call=getsubmissionsbyassignment&class=${props.class}&id=${props.id}&student=${props.student}`)
          .then(x => x.json())
          .then(x => {
            if (x.value && x.value.length > 0) {
              //这个学生有提交作业
              const submissionId = x.value[0].id;
              fetch(`api/service?call=getsubmissionoutcome&class=${props.class}&assignment=${props.id}&id=${submissionId}`).then(x => x.json()).then(x => {
                if (x.value && x.value.length > 0) {
                  let feedback, point;

                  x.value.forEach((item: any) => {
                    if (item.feedback) {
                      feedback = item.feedback.text.content;
                    }

                    if (item.points) {
                      point = item.points.points;
                    }
                  });

                  setStatus(`作业已提交,得分:${point},老师评语:${feedback ?? '无'}`);
                }
                else
                  setStatus("作业已提交，但老师还没有批改");
              });
            }
            else
              setStatus("当前还没有提交作业");
          })
      }, []);

      return (<div>{status}</div>)
    }
    useEffect(() => {
      fetch(`api/service?call=getassignmentsbyclass&id=${props.class}`).then(x => x.json()).then(x => setAssignments(x.value));
    }, [])
    return (<List items={
      assignments?.map(x => {
        return {
          key: x.id,
          header: <h3>{x.displayName}</h3>,
          content: <>
            <p>{`发布日期:${x.assignedDateTime},截止日期:${x.dueDateTime}, 总分:${x.grading.maxPoints}`}</p>
            <AssignmentStatus class={props.class} id={x.id} student={props.student} />
            <hr />
          </>
        }
      })
    }></List>)
  }

  //公告列表组件
  const PostList = (props: { class: string }) => {
    const [posts, setPosts] = useState<[any]>();//公告

    const Comment = (props: { postId: string }) => {
      const [comments, setComments] = useState<[any]>();
      useEffect(() => {
        fetch(`api/service?call=getcommentsbypost&id=${props.postId}`).then(x => x.json()).then(x => setComments(x));
      }, [])

      return (<>
        {
          comments ?
            <List
              items={
                comments?.map(x => {
                  return {
                    key: x.id,
                    header: x.comment,
                    content: `${x.author} 发表于 ${x.time}`
                  }
                })
              }></List >
            : ""}
        <hr /></>)
    }

    useEffect(() => {
      fetch(`api/service?call=getpostsbyclass&id=${props.class}`).then(x => x.json()).then(x => setPosts(x));
    }, [])
    return (<List items={
      posts?.map(x => {
        return {
          key: x.id,
          header: <h3>{x.content}</h3>,
          content: <Comment postId={x.id}></Comment>
        }
      })
    }></List>)
  }

  return (
    <Flex column fill gap="gap.medium">
      <Segment color="brand" inverted>
        <Flex column fill gap="gap.small">
          <Text content={`你好,家长`} size="larger"></Text>
          <Text content="这是家校通的家长界面，在这里你看到孩子的作业完成情况，以及老师下发的通知公告等" size="small"></Text>
        </Flex>
      </Segment>

      {parentInfo && <Segment>
        <List items={
          parentInfo.map(x => {
            return {
              key: x.class,
              header: `作为${x.student}的${x.type}, 班级编号:${x.class}`,
              content: <>
                <h2>作业列表</h2>
                <AssignmentList class={x.class} student={x.id}></AssignmentList>
                <h2>班级公告</h2>
                <PostList class={x.class}></PostList>
              </>
            }
          })
        }>
        </List>
      </Segment>}


      {!parentInfo && <Segment>
        <Form
          onSubmit={(evt) => {
            const email = (evt.target as any).txtEmail.value;
            //作为范例，这里目前没有做真正的验证
            fetch(`api/service?call=getParentInfoByEmail&email=${email}`).then(x => x.json()).then(x => {
              setParentInfo(x);
            }).then(_ => {
              localStorage.setItem("parentEmail", email);
            });
          }}
        >
          <FormInput name="txtEmail" placeholder="请使用邮箱登录"></FormInput>
          <FormButton content="登录" primary></FormButton>
        </Form>
      </Segment>}
    </Flex>
  )
}


function Student(props: { studentId: string, studentName: string, classId: string }) {

  const [settings, setSettings] = useState<[any]>();
  const [inputName, setInputName] = useState<string>();
  const [selectedType, setSelectedType] = useState<string>();

  useEffect(() => {
    fetch(`api/service?call=getParentInfoById&id=${props.studentId}`).then(x => x.json()).then(x => setSettings(x));
  }, [])

  return (
    <Flex column fill gap="gap.medium">
      <Segment color="brand" inverted>
        <Flex column fill gap="gap.small">
          <Text content={`你好,${props.studentName}`} size="larger"></Text>
          <Text content="在这里你可以设置家长信息，你需要输入邮箱地址，和选择称谓。你设置的家长，可以通过邮箱登录到家校通（网页版），然后可以查看你所在班级的作业及老师批改情况，并且可以跟你所在班级的老师进行留言互动。" size="small"></Text>
        </Flex>
      </Segment>

      {settings && settings.length > 0 &&
        <Segment>
          <h2>已经添加的家长</h2>
          <List items={settings?.map((x: any) => {
            return {
              key: x.email,
              header: x.type,
              content: x.email
            }
          })}></List>
        </Segment>
      }
      <Segment>

        <h2>添加新的家长</h2>
        <Form
          style={{ marginLeft: 20 }}
          onSubmit={() => {
            const data = {
              id: props.studentId,
              student: props.studentName,
              type: selectedType,
              email: inputName,
              class: props.classId
            };
            //调用接口提交数据
            fetch("api/service?call=addParentInfo", {
              method: "post",
              body: JSON.stringify(data)
            }).then(x => {
              let temp = settings;
              temp?.push(data);
              setSettings(temp);
            })
          }}
        >

          <FormInput
            label="家长邮箱"
            name="firstName"
            id="first-name"
            required
            defaultValue={inputName}
            onChange={(evt, data) => {
              setInputName(data?.value)
            }}
          />
          <FormDropdown
            items={['爸爸', '妈妈', '爷爷', '奶奶', '外公', '外婆', '其他']}
            search={true}
            autoSize={false}
            placeholder="选择家长称谓"
            defaultValue={selectedType}
            onChange={(evt, data) => {
              if (data.value) {
                setSelectedType(data.value as string);
              }
            }}
          />
          <FormButton content="提交" />
        </Form>
      </Segment>
    </Flex>)
}

function Teacher(props: { teacherId: string, teacherName: string, classId: string }) {
  const [posts, setPosts] = useState<[any]>();

  useEffect(() => {
    fetch(`api/service?call=getmyposts&id=${props.teacherId}`).then(x => x.json()).then(x => setPosts(x));
  }, [])

  // 评论组件
  const Comment = (props: { postId: string }) => {
    const [comments, setComments] = useState<[any]>();
    useEffect(() => {
      fetch(`api/service?call=getcommentsbypost&id=${props.postId}`).then(x => x.json()).then(x => setComments(x));
    }, [])

    return (<>
      {
        comments ?
          <List
            items={
              comments?.map(x => {
                return {
                  key: x.id,
                  header: x.comment,
                  content: `${x.author} 发表于 ${x.time}`
                }
              })
            }></List >
          : ""}
    </>)
  }

  return (
    <Flex column fill gap="gap.medium">
      <Segment color="brand" inverted>
        <Flex column fill gap="gap.small">
          <Text content={`你好,${props.teacherName}`} size="larger"></Text>
          <Text content="这是家校通的老师界面，在这里你可以发布公告，并且也可以查看家长的回复" size="small"></Text>
        </Flex>
      </Segment>

      <Segment>
        <h2>发布公告</h2>
        <Form
          style={{ marginLeft: 20 }}
          onSubmit={(evt) => {
            const text = (evt.target as any).txtcomment.value;
            fetch("api/service?call=addpost", {
              method: "post",
              body: JSON.stringify({
                teacherId: props.teacherId,
                content: text,
                class: props.classId
              })
            }).then(_ => {
              (evt.target as any).txtcomment.value = "";
            })
          }}
        >
          <FormTextArea style={{ width: 600 }} name="txtcomment" />
          <FormButton content="发布"></FormButton>
        </Form>
      </Segment>

      <Segment>
        <h2>公告列表</h2>
        <List items={posts?.map((x: any) => {
          return {
            key: x.id,
            header: <>
              <h3>{x.content}</h3>
              <p>发布于: {x.time}</p>
            </>,
            content: <Comment postId={x.id} />
          }
        })}></List>
      </Segment>
    </Flex>
  )
}


function Tab() {

  const [context, setContext] = useState<microsoftTeams.Context>();//当前Teams上下文

  useEffect(() => {
    microsoftTeams.getContext(ctx => {
      setContext(ctx);
    })
  }, [])

  return (
    <>
      {context && context.userObjectId && context.groupId && context.userPrincipalName ?
        (context.userTeamRole ?
          <Student studentId={context.userObjectId} studentName={context.userPrincipalName.split('@')[0]} classId={context.groupId} /> :
          <Teacher teacherName={context.userPrincipalName.split('@')[0]} teacherId={context.userObjectId} classId={context.groupId}></Teacher>)
        : <div>正在加载...</div>
      }
    </>
  )

}

function Configuration() {

  microsoftTeams.settings.setValidityState(true);
  microsoftTeams.settings.registerOnSaveHandler(evt => {
    microsoftTeams.settings.setSettings({
      entityId: "eduapp",//这个id，在真实的项目开发中，需要考虑设计成一个唯一性的字段，因为用户可以在任何团队和频道添加你这个选项卡，即便是同一个频道，也可以添加多次，那么如何进行区分呢？这个id 很重要。
      websiteUrl: "https://2adeb3f8ee53.ngrok.io",//在Teams中的选项卡，除了可以在Teams中打开外，还可以在浏览器中打开。如果在浏览器中打开，会调用这个地址。
      contentUrl: "https://2adeb3f8ee53.ngrok.io/tab",//这是在Teams里面打开的地址。如果要区分不同的选项卡实例，可以考虑在这个地址后面添加参数。
      suggestedDisplayName: "家校通"//这是推荐给用户的选项卡名称
    });


    evt.notifySuccess();
  })

  return (
    <>
      <h1>家校通应用</h1>
      <p>这个页面用来让用户进行配置，通常来说，这里会有一些可以配置的选项，然后根据选项的值，决定用户是否可以点击“保存”按钮，点击“保存”按钮时，会调用有关的接口，进行保存操作，并且通知客户端关闭当前的配置窗口。</p>
    </>
  )
}

function App() {
  microsoftTeams.initialize();

  return (
    <Provider theme={teamsTheme}>
      <Router>
        <Route exact path="/" component={Parent}></Route>
        <Route exact path="/config" component={Configuration}></Route>
        <Route exact path="/tab" component={Tab}></Route>
      </Router>
    </Provider>
  );
}

export default App;

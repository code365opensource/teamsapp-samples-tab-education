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
      const [status, setStatus] = useState<string>("Checking...");

      useEffect(() => {
        fetch(`api/service?call=getsubmissionsbyassignment&class=${props.class}&id=${props.id}&student=${props.student}`)
          .then(x => x.json())
          .then(x => {
            if (x.value && x.value.length > 0) {
              //这个学生有提交作业
              const submissionId = x.value[0].id;
              fetch(`api/service?call=getsubmissionoutcome&class=${props.class}&assignment=${props.id}&id=${submissionId}`).then(x => x.json()).then(x => {
                let feedback, point;
                x.value.forEach((item: any) => {
                  if (item.feedback) {
                    feedback = item.feedback.text.content;
                  }

                  if (item.points) {
                    point = item.points.points;
                  }
                });

                if (feedback || point)
                  setStatus(`Submission submitted, score is: ${point ?? 0}, feedback from teacher is: ${feedback ?? 'None'}`);
                else
                  setStatus(`Submission submitted, but did't get feedback from teacher yet.`);
              });
            }
            else
              setStatus("Need to submit the submission soon.");
          })
      }, []);

      return (<div style={{ color: status.startsWith("Need") ? "red" : status.endsWith('yet.') ? "blue" : "green" }}>{status}</div>)
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
            <p>{`Assigned date: ${x.assignedDateTime}, Due date: ${x.dueDateTime}, Points: ${x.grading.maxPoints}`}</p>
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
                    content: `${x.author} published at ${x.time}`
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
          <Text content={`Hi,Parents`} size="larger"></Text>
          <Text content="Welcome to the Family Engagement system, as a parent, you can view the basic information of your children's assignment and status, you can also view the announments in one place, communicate with teachers by your convinence way." size="small"></Text>
        </Flex>
      </Segment>

      {parentInfo && <Segment>
        <List items={
          parentInfo.map(x => {
            return {
              key: x.class,
              header: `As ${x.student}'s ${x.type}, Class number is: ${x.class}`,
              content: <>
                <h2>Assignments</h2>
                <AssignmentList class={x.class} student={x.id}></AssignmentList>
                <h2>Announments</h2>
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
          <FormInput name="txtEmail" placeholder="Use your email to login"></FormInput>
          <FormButton content="Login" primary></FormButton>
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
          <Text content={`Hi,${props.studentName}`} size="larger"></Text>
          <Text content="As a student, you can add your parents into the Family Engagement system, so that they can view the basic information of your assignment and the submission status, they also can communicate with your teachers smoothly. Everything is under your control, you can disable the settings at anytime if you want. " size="small"></Text>
        </Flex>
      </Segment>

      {settings && settings.length > 0 &&
        <Segment>
          <h2>Added parents</h2>
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

        <h2>Add your parent</h2>
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
            label="Email"
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
            placeholder="Select Type"
            defaultValue={selectedType}
            onChange={(evt, data) => {
              if (data.value) {
                setSelectedType(data.value as string);
              }
            }}
          />
          <FormButton content="Submit" />
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
                  content: `${x.author} published at ${x.time}`
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
          <Text content={`Hi,${props.teacherName}`} size="larger"></Text>
          <Text content="Welcome to the Family Engagement system, As a teacher, you can post announments here, and communicate with parents in one place." size="small"></Text>
        </Flex>
      </Segment>

      <Segment>
        <h2>Post Announment</h2>
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
          <FormButton content="Publish"></FormButton>
        </Form>
      </Segment>

      <Segment>
        <h2>Announments</h2>
        <List items={posts?.map((x: any) => {
          return {
            key: x.id,
            header: <>
              <h3>{x.content}</h3>
              <p>Published at: {x.time}</p>
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
        : <div>Loading...</div>
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

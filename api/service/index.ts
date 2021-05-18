import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import { v4 } from "uuid";

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    const FileSync = require("lowdb/adapters/FileSync");
    const lowdb = require("lowdb");
    const adapter = new FileSync("./data.json");
    const db = lowdb(adapter);

    db.defaults({ studentSettings: [], posts: [], comments: [] }).write();

    //这个接口用来模拟服务调用
    let responseMessage = {};
    // 学生相关
    //1. 读取家长信息（姓名，邮箱，称谓） GET, getParentInfoById
    if (req.method.toLowerCase() === "get" && req.query.call === "getParentInfoById" && req.query.id) {
        responseMessage = db.get("studentSettings").filter({ id: req.query.id }).value()
    }

    //2. 保存家长信息（姓名，邮箱，称谓） POST,addParentInfo
    else if (req.method.toLowerCase() === "post" && req.query.call === "addParentInfo") {
        db.get("studentSettings").push(JSON.parse(req.rawBody)).write();
        responseMessage = {
            body: "保存成功"
        }
    }

    // 老师相关
    //1. 发送公告消息
    else if (req.method.toLowerCase() === "post" && req.query.call === "addpost") {
        const data = JSON.parse(req.rawBody);
        db.get("posts").push({
            id: v4(),
            teacherId: data.teacherId,
            content: data.content,
            time: new Date().toLocaleString()
        }).write();

    }
    //2. 读取某个老师发布的公告列表
    else if (req.method.toLowerCase() === "get" && req.query.call === "getmyposts" && req.query.id) {
        responseMessage = db.get("posts").filter({ teacherId: req.query.id }).value();
    }

    //3. 获取某个公告的评论
    else if (req.method.toLowerCase() === "get" && req.query.call === "getcommentsbypost" && req.query.id) {
        responseMessage = db.get("comments").filter({ postId: req.query.id }).value();
    }

    // 家长相关
    // 1. 根据email获取孩子和班级信息
    else if (req.method.toLowerCase() === "get" && req.query.call === "getParentInfoByEmail" && req.query.email) {
        responseMessage = db.get("studentSettings").filter({ email: req.query.email }).value()
    }

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;
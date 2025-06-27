### 任务信息

| 公司名称  | XX公司                                                |
|:-----:| --------------------------------------------------- |
| 系统名称  | XXX管理系统                                             |
| 漏洞名称  | 未授权访问                                               |
| 域名    | www.test.com                                        |
| 漏洞URL | https://www.test.com:8080/admin.spring?encode=utf-8 |
| 危害级别  | 中危                                                  |
| 日期    | 2025年4月10号                                          |
| 测试人员  | Nika                                                |

### 未授权访问

```
GET /admin.spring?encode=utf-8 HTTP/1.1
Host: www.test.com
Accept: */*
Accept-Language: zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2
Accept-Encoding: gzip, deflate, br
X_requested_with: XMLHttpRequest
Content-Type: multipart/form-data; boundary=----geckoformboundary8de296bfe899025ea066c083c9f6aa3f
Content-Length: 926
Sec-Fetch-Dest: empty
Sec-Fetch-Mode: cors
Sec-Fetch-Site: same-origin
Te: trailers
Connection: keep-alive
```

1.未授权访问

![](C:\Users\wl\Pictures\desktop.jpg)

2.未授权操作

<img src="file:///C:/Users/wl/Pictures/desktop.jpg" title="" alt="" width="668">

---

### 笔记

不会被输出(will not print)

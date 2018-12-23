<%--
  Created by IntelliJ IDEA.
  User: Administrator
  Date: 2018/07/07
  Time: 下午 12:12
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>

<html>
<head>
    <title>欢迎</title>
    <script src="https://cdn.staticfile.org/jquery/1.10.2/jquery.min.js">
    </script>
</head>
<body>
    <li>
        <span>上&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;传：</span>
        <span class="input">
       <input type="file" id="upfile" name="upfile" placeholder=""/>
    </span>
        <button onclick="importExp();">导入</button>
        <span>格式：.xls</span>
    </li>
</body>

<script>
    //导入文件
    function importExp() {
        var formData = new FormData();
        var name = $("#upfile").val();
        formData.append("file",$("#upfile")[0].files[0]);
        formData.append("name",name);
        $.ajax({
            url : '/importExp',
            type : 'POST',
            async : false,
            data : formData,
            // 告诉jQuery不要去处理发送的数据
            processData : false,
            // 告诉jQuery不要去设置Content-Type请求头
            contentType : false,
            beforeSend:function(){
                console.log("正在进行，请稍候");
            },
            success : function(responseStr) {
                if(responseStr=="01"){
                    alert("导入成功");
                }else{
                    alert("导入失败");
                }
            }
        });
    }
</script>
</html>

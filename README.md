# loop_request_write_excel
循环发送请求获取网页参数后格式化，然后将结果写到本地excel内

# 依赖安装
```shell
pip install openpyxl
pip install requests
```

# 使用注意
- 请求地址、请求头、请求参数、数据格式化需要根据不同的场景进行替换
- 尤其注意请求头内一般会加一些防爬策略，本示例中cookie值有效期较短，需要手动获取替换后再执行


# 声明
- 本仓库代码仅供个人学习使用，请不要使用此工具搞破坏
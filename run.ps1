# 将 PowerShell 执行策略设置为 Unrestricted 以执行脚本
# 运行 Python 脚本 test.py，参数从0到100
0..99 | ForEach-Object {
    python crawler.py $_
}
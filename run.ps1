# 将 PowerShell 执行策略设置为 Unrestricted 以执行脚本
0..99 | ForEach-Object {
    python crawler.py $_
}
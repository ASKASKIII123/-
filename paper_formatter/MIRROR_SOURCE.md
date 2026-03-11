# 镜像源使用说明

## 🚀 默认配置

本软件已配置使用 **清华大学镜像源** 进行加速下载，无需额外配置。

## 📦 自动使用的镜像源

### 启动器自动配置

当你运行 `启动软件.bat` 或 `launcher.py` 时，系统会自动使用以下镜像源：

```bash
https://pypi.tuna.tsinghua.edu.cn/simple
```

## 🔧 手动使用镜像源

如果需要手动安装依赖包，可以使用以下命令：

### 方法1：使用清华镜像源（推荐）

```bash
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
```

### 方法2：使用阿里云镜像源

```bash
python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
```

### 方法3：使用豆瓣镜像源

```bash
python -m pip install -r requirements.txt -i https://pypi.douban.com/simple/ --trusted-host pypi.douban.com
```

### 方法4：使用华为云镜像源

```bash
python -m pip install -r requirements.txt -i https://mirrors.huaweicloud.com/repository/pypi/simple/ --trusted-host mirrors.huaweicloud.com
```

## 🌍 国内常用镜像源列表

| 镜像源 | 地址 | 状态 |
|--------|------|------|
| 清华大学 | https://pypi.tuna.tsinghua.edu.cn/simple | ⭐ 推荐 |
| 阿里云 | https://mirrors.aliyun.com/pypi/simple/ | 稳定 |
| 豆瓣 | https://pypi.douban.com/simple/ | 稳定 |
| 华为云 | https://mirrors.huaweicloud.com/repository/pypi/simple/ | 稳定 |
| 中科大 | https://pypi.mirrors.ustc.edu.cn/simple/ | 稳定 |
| 腾讯云 | https://mirrors.cloud.tencent.com/pypi/simple/ | 稳定 |

## 🔍 永久配置镜像源

### Windows系统

1. 创建或编辑 `%APPDATA%\pip\pip.ini` 文件
2. 添加以下内容：

```ini
[global]
index-url = https://pypi.tuna.tsinghua.edu.cn/simple
trusted-host = pypi.tuna.tsinghua.edu.cn
```

### Linux/Mac系统

1. 创建或编辑 `~/.pip/pip.conf` 文件
2. 添加以下内容：

```ini
[global]
index-url = https://pypi.tuna.tsinghua.edu.cn/simple
trusted-host = pypi.tuna.tsinghua.edu.cn
```

## ⚡ 临时使用镜像源

### 安装单个包

```bash
python -m pip install 包名 -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
```

### 升级pip

```bash
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
```

## 📊 镜像源速度对比

根据网络环境不同，各镜像源的速度可能有所差异。建议：

1. **首选**：清华大学镜像源（本软件默认）
2. **备选**：阿里云镜像源
3. **测试**：如果速度慢，可以尝试其他镜像源

## 🔧 切换镜像源

如果需要切换到其他镜像源，修改以下文件：

### 修改 launcher.py

找到第56-58行，修改镜像源地址：

```python
subprocess.check_call([
    sys.executable, "-m", "pip", "install", "-r", "requirements.txt",
    "-i", "https://mirrors.aliyun.com/pypi/simple/",  # 修改这里
    "--trusted-host", "mirrors.aliyun.com"            # 修改这里
])
```

### 修改 run_enhanced.bat

找到第41-42行，修改镜像源地址：

```batch
python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
```

## ❓ 常见问题

### Q: 为什么使用镜像源？

A: 国内访问官方PyPI速度较慢，使用国内镜像源可以大幅提升下载速度。

### Q: 镜像源安全吗？

A: 国内主流镜像源都是安全的，它们会定期同步官方PyPI的内容。

### Q: 镜像源更新及时吗？

A: 主流镜像源通常每10-30分钟同步一次官方PyPI，基本可以满足需求。

### Q: 如果镜像源无法访问怎么办？

A: 可以尝试其他镜像源，或者暂时使用官方源（去掉 `-i` 参数）。

## 📝 注意事项

1. **首次安装**：建议使用镜像源加速
2. **更新包**：同样可以使用镜像源
3. **网络问题**：如果镜像源无法访问，尝试切换其他镜像源
4. **SSL证书**：使用 `--trusted-host` 参数避免SSL证书问题

## 🎯 推荐配置

对于大多数用户，推荐使用默认的清华大学镜像源：

```bash
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
```

这个配置已经在软件的启动器中自动应用，无需手动设置！
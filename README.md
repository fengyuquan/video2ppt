# 从视频中提取关键帧并创建PowerPoint演示文稿

## 项目简介

本项目旨在使用Python从视频中提取关键帧，并将这些关键帧自动生成PowerPoint演示文稿。该项目使用OpenCV库进行视频处理和关键帧提取，使用python-pptx库创建PowerPoint演示文稿。

## 目录

- [项目简介](#项目简介)
- [目录](#目录)
- [环境配置](#环境配置)
- [安装依赖](#安装依赖)
- [使用方法](#使用方法)

## 环境配置

### 使用 `venv` 创建虚拟环境

1. 创建虚拟环境：
    ```bash
    python -m venv myenv
    ```

2. 激活虚拟环境：
    - **Windows**:
      ```bash
      myenv\Scripts\activate
      ```
    - **macOS和Linux**:
      ```bash
      source myenv/bin/activate
      ```

### 设置pip国内源（可选）

为了加速依赖包的下载速度，可以设置pip的国内源：

```bash
pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

## 安装依赖

在激活的虚拟环境中，运行以下命令安装所需的依赖包：

```bash
pip install opencv-python python-pptx
```

## 使用方法

假设你的视频文件名为 `sample_video.mp4`，并且你希望生成的PowerPoint文件名为 `output_presentation.pptx`，你可以按照以下步骤操作：

1. 将 `sample_video.mp4` 放置在项目目录中。
2. 修改脚本中的 `video_path` 和 `output_ppt_path`：
    ```python
    video_path = "sample_video.mp4"
    output_ppt_path = "output_presentation.pptx"
    ```
3. 运行脚本：
    ```bash
    python script_name.py
    ```

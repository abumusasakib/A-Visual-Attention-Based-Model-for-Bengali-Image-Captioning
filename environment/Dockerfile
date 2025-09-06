# hash:sha256:ec4087bcc7fccb5e7b0fcb7aab47ef6ddc1a00d9c7fc571dce23663271dd3457
FROM registry.codeocean.com/codeocean/miniconda3:4.9.2-cuda11.7.0-cudnn8-ubuntu20.04

ARG DEBIAN_FRONTEND=noninteractive

RUN sed -i 's|http://archive.ubuntu.com/ubuntu/|http://mirrors.edge.kernel.org/ubuntu/|g' /etc/apt/sources.list \
    && sed -i 's|http://security.ubuntu.com/ubuntu/|http://mirrors.edge.kernel.org/ubuntu/|g' /etc/apt/sources.list

RUN apt-get update && apt-get install -y --fix-missing --no-install-recommends \
    build-essential fontconfig fonts-dejavu-core fonts-freefont-ttf gcc \
    libcairo2-dev libfribidi-dev libgl1 libglib2.0-0 libharfbuzz-dev \
    libpango1.0-dev libpangocairo-1.0-0 pkg-config python3-dev wget \
    && rm -rf /var/lib/apt/lists/*

RUN apt-get update && apt-get install -y --fix-missing --no-install-recommends uuid-dev \
    && rm -rf /var/lib/apt/lists/*


RUN pip3 install -U --no-cache-dir \
    jupyter==1.1.1 \
    matplotlib==3.7.5 \
    mplcairo==0.6.1 \
    numpy==1.24.4 \
    opencv-python==4.12.0.88 \
    pangocairocffi==0.7.0 \
    pillow==10.4.0 \
    protobuf==3.20.0 \
    pycairo==1.26.1 \
    scikit-learn==1.3.2 \
    tensorflow==2.8.0 \
    tqdm==4.67.1

COPY postInstall /
RUN /postInstall

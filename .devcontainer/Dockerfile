FROM python:3.10.14
USER root

RUN apt-get update && \
    apt-get -y install --reinstall ca-certificates && \
    apt-get -y install software-properties-common && \
    apt-get -y install iputils-ping net-tools htop && \
    apt-get -y install fonts-ipafont && \
    apt-get -y install tree && \
    apt-get -y install python3-opencv && \
    apt-get -y install graphviz libgraphviz-dev pkg-config && \
    pip install --upgrade pip

# COPY requirements.txt /tmp/requirements.txt
# RUN pip install -r /tmp/requirements.txt

# Install Basic Modules
RUN pip install scikit-learn ipykernel jupyter numpy pandas matplotlib japanize-matplotlib seaborn pillow python-dotenv tqdm

# Install OpenCV Modules
RUN pip install opencv-python

# Install Python-pptx
RUN pip install python-pptx

# Install LLM Modules
RUN pip install langchain-core==0.3.0
RUN pip install langchain-community==0.3.0
RUN pip install langchain-openai==0.2.0
RUN pip install langgraph==0.2.22
RUN pip install langchain-experimental==0.3.0
RUN pip install pygraphviz
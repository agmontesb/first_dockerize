# Use a Python base image with Tkinter support
FROM python:3.11-slim-bookworm

# Install necessary X11 client libraries and Tkinter dependencies
RUN apt-get update && apt-get install -y \
    libxext6 \
    libxrender1 \
    libxtst6 \
    libxi6 \
    tk \
    && rm -rf /var/lib/apt/lists/*

# Set the working directory in the container
WORKDIR /app
# Copy your application files into the container
COPY ./src /app
# Install any Python dependencies (if you have a requirements.txt)
# RUN pip install -r requirements.txt
# Set the DISPLAY environment variable to connect to the X server on the host
ENV DISPLAY=host.docker.internal:0.0
# Command to run your application
CMD ["python", "worksheetui.py"]

# To run the resulting image:
# 1 - In a windows powershell: 
#     a - run wsl --shutdown. This stop the WSL2 instance.
#     b - run wsl. This will start a new WSL2 instance.
#     c - run echo $DISPLAY. This will a ":0" or "0:0" value if the X server is running.
# 2 - Run the following docker command:
#     docker run -it --rm -e DISPLAY=$DISPLAY -v /tmp/.X11-unix:/tmp/.X11-unix worksheetui:v2
#     Explanation of the docker run options:
#        - docker run -it --rm: Runs the container interactively (-i), allocates a pseudo-TTY (-t), and automatically removes it when the app exits (--rm).
#          -e DISPLAY=$DISPLAY: Passes the DISPLAY environment variable from your active WSL session directly into the container. This is the critical step for X11 forwarding.
#          -v /tmp/.X11-unix:/tmp/.X11-unix: Mounts the X11 socket as a volume, which allows the container to communicate with the X server on your Windows host.
#          worksheetui:v2: The name of the image you built. 
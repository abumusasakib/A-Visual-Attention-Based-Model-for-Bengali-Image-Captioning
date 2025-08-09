# Model Training Capsule Built using Codeocean

## Project Building and Running

### Build image

```cmd
cd environment && docker build . --tag aed8529a-ded6-4419-ad11-4a8ce0dd580b
```

### Run image

```cmd
docker run --platform linux/amd64 --rm --gpus all --workdir /code --volume "%cd%/data":/data --volume "%cd%/code":/code --volume "%cd%/results":/results aed8529a-ded6-4419-ad11-4a8ce0dd580b bash run
```

## Running Jupyter for Development

1. **Start a Docker Container:**
   You can start a new container from the image you built using the following command in Command Line:

   ```cmd
   docker run -p 8888:8888 -it --platform linux/amd64 --rm --gpus all --workdir /code --volume "%cd%/data":/data --volume "%cd%/code":/code --volume "%cd%/results":/results aed8529a-ded6-4419-ad11-4a8ce0dd580b /bin/bash
   ```

   This command will start a new container based on the image tagged as `aed8529a-ded6-4419-ad11-4a8ce0dd580b` and open an interactive shell (`/bin/bash`) within the container.

2. **Activate Miniconda Environment:**
   Before launching the Jupyter Notebook server, ensure that you have activated your Miniconda environment. If you haven't activated it yet, you can do so by running:

   ```bash
   source /opt/conda/bin/activate
   ```

   This command activates the Miniconda environment.

3. **Launch Jupyter Notebook:**
   Once your Miniconda environment is activated, you can launch the Jupyter Notebook server in Command Line by running:

   ```bash
   jupyter notebook --ip=0.0.0.0 --port=8888 --no-browser --allow-root
   ```

   This command will start the Jupyter Notebook server and open a web browser with the Jupyter Dashboard, where you can navigate your files and create or open notebooks.

4. **Access Jupyter Notebook:**
   After running the `jupyter notebook` command, you should see output in your terminal with a URL that starts with `http://127.0.0.1:8888`. Open this URL in your web browser, and you should be directed to the Jupyter Dashboard, where you can create or open notebooks.

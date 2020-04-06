FROM continuumio/miniconda3

WORKDIR /usr/src/app
COPY . .

RUN conda update -n base -c defaults conda 
RUN while read requirement; do conda install -c conda-forge --yes $requirement || pip install $requirement; done < requirements.txt

CMD [ "python", "main.py" ]
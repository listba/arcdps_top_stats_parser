# using ubuntu LTS version
FROM ubuntu:22.04 AS builder-image
ENV PYTHON_VS=3.11
RUN apt-get update && apt-get install --no-install-recommends -y "python${PYTHON_VS}" "python${PYTHON_VS}-dev" "python${PYTHON_VS}-venv" python3-pip python3-wheel build-essential && \
   apt-get clean && rm -rf /var/lib/apt/lists/*

# create and activate virtual environment
RUN "python${PYTHON_VS}" -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# install requirements
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

FROM ubuntu:22.04 AS runner-image
ENV PYTHON_VS=3.11
RUN apt-get update && apt-get install --no-install-recommends -y "python${PYTHON_VS}" python3-venv && \
   apt-get clean && rm -rf /var/lib/apt/lists/*

COPY --from=builder-image /opt/venv /opt/venv

WORKDIR /app
COPY . /app


# activate virtual environment
ENV VIRTUAL_ENV=/opt/venv
ENV PATH="/opt/venv/bin:$PATH"


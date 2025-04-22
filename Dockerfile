# Use Python slim image from Amazon ECR Public
FROM public.ecr.aws/docker/library/python:3.9-slim

# Set working directory
WORKDIR /app

# Ensure output is logged immediately
ENV PYTHONUNBUFFERED=1

# Install dependencies
COPY requirements.txt .
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Default command to run the application
CMD ["python", "app.py"]

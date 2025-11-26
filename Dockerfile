# Use an official lightweight Python image
FROM python:3.11-slim

# Set the working directory inside the container
WORKDIR /app

# Copy only the requirements file first to leverage Docker cache
COPY server/requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the server-side code (including the 'app' module and 'templates')
COPY server/ /app/server/

# Copy the web (static files) directory
COPY web/ /app/web/

# Expose the port the app runs on
EXPOSE 8000

# Command to run the application, mimicking your local command
# We use 'server.app.main:app' because we are in the /app (root) folder
CMD ["python", "-m", "uvicorn", "server.app.main:app", "--host", "0.0.0.0", "--port", "8000"]
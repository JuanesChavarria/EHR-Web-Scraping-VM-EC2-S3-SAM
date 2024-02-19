# EHR-Web-Scraping-VM-EC2-S3-SAM

## Overview

This project is aimed at deploying a virtual machine in Amazon EC2 to execute web scraping algorithms for extracting patient information from an Electronic Health Record (EHR) website. The gathered data is then batch-synced into a Customer Relationship Management (CRM) system for our healthcare company.

## Features

- **Web Scraping:** Utilizes web scraping algorithms to extract patient information.
- **Security Measures:** Implements random events to bypass security measures against bots on the EHR website.
- **AWS Integration:** Deploys virtual machines in EC2 and utilizes S3 to generate download links for the extracted files.

## Architecture

The project is designed to work within the Serverless Application Model (SAM) architecture. The `template.yaml` file and other essential files for SAM are not revealed for security purposes.

## Testing

The application has been tested using Docker and containers. You can test the application by running the endpoint created within `app.py`.

## Deployment

To deploy the project, follow these steps:

1. Clone the repository.
2. Deploy the SAM application using the a `template.yaml`.

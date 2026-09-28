FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app
COPY pyproject.toml README.md /app/
COPY presentation_builder /app/presentation_builder
COPY prompts /app/prompts
RUN pip install --no-cache-dir .

RUN useradd --create-home --uid 10001 appuser \
    && mkdir -p /app/service_data \
    && chown -R appuser:appuser /app
USER appuser

EXPOSE 8000
CMD ["presentation-service"]


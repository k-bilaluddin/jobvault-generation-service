// OTel must be initialised before any other imports that touch instrumented modules
import { startTelemetry, shutdownTelemetry } from './telemetry';
startTelemetry();

import { startServer } from './rest/server';

const PORT = parseInt(process.env.PORT ?? '3000', 10);
startServer(PORT);

const shutdown = async () => {
  await shutdownTelemetry();
  process.exit(0);
};

process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);

import express from 'express';
import router from './router';

export function createServer() {
  const app = express();
  app.use(express.json({ limit: '2mb' }));
  app.use(router);
  return app;
}

export function startServer(port: number): void {
  const app = createServer();
  app.listen(port, () => {
    console.log(`jobvault-generation-service listening on port ${port}`);
  });
}

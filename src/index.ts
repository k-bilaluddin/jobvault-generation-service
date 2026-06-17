import { startServer } from './rest/server';

const PORT = parseInt(process.env.PORT ?? '3000', 10);
startServer(PORT);

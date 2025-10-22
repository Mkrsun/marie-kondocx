import pino from "pino";

export function createLogger(level: "debug" | "info" | "warn" | "error" = "info") {
  return pino({
    level,
    transport: {
      target: "pino-pretty",
      options: {
        colorize: true,
        translateTime: "HH:MM:ss",
        ignore: "pid,hostname"
      }
    }
  });
}

/**
 * Crea un logger que escribe tanto a consola como a archivo
 */
export function createFileLogger(
  level: "debug" | "info" | "warn" | "error" = "info",
  logFilePath?: string
) {
  const targets = [
    {
      target: "pino-pretty",
      options: {
        colorize: true,
        translateTime: "HH:MM:ss",
        ignore: "pid,hostname"
      },
      level
    }
  ];

  // Si se especifica archivo de log, agregarlo como destino
  if (logFilePath) {
    targets.push({
      target: "pino/file",
      options: {
        destination: logFilePath
      },
      level
    } as any);
  }

  return pino({
    level,
    transport: {
      targets
    }
  });
}

export type Logger = ReturnType<typeof createLogger>;

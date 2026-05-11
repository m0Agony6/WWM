export interface AppConfig {
  reportPrefixes: { [key: string]: string };
  supervisors: { [key: string]: string };
  moldPrefixes: { [key: string]: string };
}

export interface ChatMessage {
  role: 'user' | 'ai';
  text: string;
}

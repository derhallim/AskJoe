// Import required packages
import * as restify from "restify";
import { Application, PromptManager, ActionPlanner , TurnState, OpenAIModerator, OpenAIModel, AI, ApplicationBuilder} from '@microsoft/teams-ai'
import path from 'path'
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  MemoryStorage, 
  ActivityTypes
} from "botbuilder";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";
import config from "./config";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
  if (context.activity.type === "message") {
    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
      "OnTurnError Trace",
      `${error}`,
      "https://www.botframework.com/schemas/error",
      "TurnError"
    );

    // Send a message to the user
    await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create the bot that will handle incoming messages.
const bot = new TeamsBot();

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

interface ConversationState {}
type ApplicationTurnState = TurnState<ConversationState>;


const model = new OpenAIModel({
  // OpenAI Support
  apiKey: process.env.OPENAI_KEY!,
  defaultModel: 'gpt-35-turbo-16k',

  // Azure OpenAI Support
  azureApiKey: process.env.AZURE_OPENAI_KEY!,
  azureDefaultDeployment: 'gpt-35-turbo-16k',
  azureEndpoint: process.env.AZURE_OPENAI_ENDPOINT!,
  azureApiVersion: '2023-03-15-preview',

  // Request logging
  logRequests: true
});

const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, '../prompts')
});

const storage = new MemoryStorage();

const planner = new ActionPlanner({
  model,
  prompts,
  defaultPrompt: 'chat'
});

const app = new Application<ApplicationTurnState>({
  storage,
  ai: {
      planner
  }
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await app.run(context);
  });
});

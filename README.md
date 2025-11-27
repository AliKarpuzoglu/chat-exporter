# Chat Exporter

Export your Microsoft Teams chat history directly from your browser. No admin permissions needed, no servers involved, just a simple tool that saves your conversations to JSON or plain text.

**[Install from Chrome Web Store](https://chrome.google.com/webstore/detail/mjmhnelnamahnpjcpnnabbhgfneggebk)**

If you can read a chat in Teams, you should be able to save it. This extension just automates the tedious work of scrolling and copying everything manually.

## What it does

Adds an export button to your Teams chat view. Click it, and the extension will:
- Auto-scroll through the conversation to load the full history
- Extract message text, sender names, timestamps, and IDs
- Export as JSON or TXT (configurable in settings)

## Installation options

### Option 1: Chrome Web Store
Install directly from the [Chrome Web Store](https://chrome.google.com/webstore/detail/mjmhnelnamahnpjcpnnabbhgfneggebk).

### Option 2: Load from GitHub
1. Clone or download this repo
2. Go to `chrome://extensions`
3. Enable "Developer mode"
4. Click "Load unpacked" and select the folder

### Option 3: Console script (no extension needed)
If you'd rather not install anything, copy the contents of [`console-script.js`](console-script.js) and paste it into your browser console while on a Teams chat page. Then run:

```js
exportTeamsChat()        // exports as JSON
exportTeamsChat("txt")   // exports as plain text
```

## Important disclaimer

All this tool does is scroll up and grab what's visible in the DOM. I make no guarantees about accuracy or completeness. Messages might be missing, formatting could be off, or things might break if Teams changes their interface.

**For critical conversations, consider recording a video of the chat as a backup.** Screen recordings are more reliable than any automated export tool.

## If it stops working

Teams might change their interface, which can break this tool. If you run into issues, please open an issue on GitHub and I'll see if I can get it working again. No promises, but I'll do what I can.

## Security and transparency

I get it. Letting an extension access your work chats is a legitimate security concern. That's why everything is open:

You have three ways to verify and use this tool:
1. **Review the source** - The full code is right here in this repo
2. **Run the script directly** - Use `console-script.js` in your browser console without installing anything
3. **Check the extension** - Click the settings gear and "View Source Code" to see exactly what's running

Review the code yourself before trusting it with your data.

## Privacy

Everything runs locally in your browser. No data gets sent anywhere. No tracking, no analytics, no servers. Your messages stay on your machine unless you decide to move them yourself.

## When this is useful

- Backing up important conversations before you send them to your favorite LLM completely ignoring confidentiality
- Keeping records of project discussions
- Preserving messages that matter to you
- Documentation or compliance purposes

## Limitations

- Only works with Teams Web (not the desktop app)
- Only exports what's already visible to you
- Your organization's data policies still apply. Make sure you're following them.

This tool isn't affiliated with Microsoft. It's just a side project to solve a common problem: making it easier to keep your own chat history. And I can't believe Teams does not have that functionality either.

---

If this tool helped you, I'd love it if you follow me on [BlueSky](https://bsky.app/profile/watermelonson.bsky.social). I'll share other projects there.

Or [buy me a coffee](https://buymeacoffee.com/watermelonson).

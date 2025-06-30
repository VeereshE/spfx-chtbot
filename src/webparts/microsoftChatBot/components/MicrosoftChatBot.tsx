import * as React from 'react';
import { useState } from 'react';
import styles from './MicrosoftChatBot.module.scss';
import type { IMicrosoftChatBotProps } from './IMicrosoftChatBotProps';
import { FaPaperPlane, FaRobot, FaTimes } from 'react-icons/fa';
import { MSGraphClientV3 } from '@microsoft/sp-http';

interface IMessage {
  from: 'bot' | 'user';
  text: string;
  time: string;
}

const MicrosoftChatBot: React.FC<IMicrosoftChatBotProps> = ({ hasTeamsContext, userDisplayName, context }) => {
  const [inputValue, setInputValue] = useState<string>('');
  const [messages, setMessages] = useState<IMessage[]>([
    {
      from: 'bot',
      text: 'Hi, how are you?',
      time: '2/6/2022 10:45 AM'
    },
    {
      from: 'user',
      text: "I'm good, thanks! How about you?",
      time: '2/6/2022 10:46 AM'
    },
    {
      from: 'bot',
      text: 'I am fine, how can I assist you?',
      time: '2/6/2022 10:48 AM'
    }
  ]);
  const [isOpen, setIsOpen] = useState<boolean>(false);

  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setInputValue(event.target.value);
  };

  const handleSend = async () => {
    if (!inputValue.trim()) return;

    const newMessage: IMessage = {
      from: 'user',
      text: inputValue.trim(),
      time: new Date().toLocaleString()
    };

    setMessages(prevMessages => [...prevMessages, newMessage]);
    setInputValue('');

    try {
      const graphClient: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');

      const recipientEmail = 'yashashree.kommajoshyula@rsmus.com'; 
      const chatResponse = await graphClient.api('/chats')
        .post({
          chatType: 'oneOnOne',
          members: [
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${recipientEmail}`
              //`https://graph.microsoft.com/v1.0/users/${yashashree.kommajoshyula@rsmus.com}`
            },
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/me`
            }
          ]
        });

      await graphClient.api(`/chats/${chatResponse.id}/messages`)
        .post({
          body: {
            content: newMessage.text
          }
        });
       console.log("Message Added ..............")
    } catch (error) {
     console.error('Error sending message via Graph API:', error);
      
    }
  };

  const handleKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter') {
      handleSend();
    }
  };

  const toggleChatBot = () => {
    setIsOpen(prev => !prev);
  };

  return (
    <section className={styles.chatWrapper}>
      {!isOpen && (
        <div className={styles.avatarIcon} onClick={toggleChatBot} aria-label="Open chat">
          <FaRobot size={32} />
        </div>
      )}

      {isOpen && (
        <div className={`${styles.microsoftChatBot} ${hasTeamsContext ? styles.teams : ''}`}>
          <div className={styles.header}>
            <img src="https://rsmus.com/content/experience-fragments/rsm/us/en/site/header/master/_jcr_content/root/globalheader/mainnav/logo.coreimg.png/1648142668633/logo.png" alt="Company Logo" className={styles.logo} />
            <div>
              <h1>Welcome {userDisplayName}!</h1>
              <h3>RSM Support Center</h3>
            </div>
            <div className={styles.closeIcon} onClick={toggleChatBot} aria-label="Close chat">
              <FaTimes size={32} />
            </div>
          </div>

          <div className={styles.chatContainer}>
            {messages.map((msg, index) => (
              <div
                key={index}
                className={`${styles.chatMessage} ${msg.from === 'bot' ? styles.botMessage : styles.userMessage}`}
              >
                <div className={styles.messageHeader}>
                  <span className={styles.userName}>
                    {msg.from === 'bot' ? 'Bot' : userDisplayName}
                  </span>
                  <span className={styles.timeStamp}>{msg.time}</span>
                </div>
                <p className={styles.messageText}>{msg.text}</p>
              </div>
            ))}
          </div>

          <div className={styles.inputContainer}>
            <input
              type="text"
              placeholder="Type a message"
              value={inputValue}
              onChange={handleInputChange}
              onKeyDown={handleKeyPress}
              className={styles.inputBox}
              aria-label="Message input"
            />
            <button
              onClick={handleSend}
              className={styles.sendButton}
              aria-label="Send message"
              type="button"
            >
              <FaPaperPlane />
            </button>
          </div>
        </div>
      )}
    </section>
  );
};

export default MicrosoftChatBot;

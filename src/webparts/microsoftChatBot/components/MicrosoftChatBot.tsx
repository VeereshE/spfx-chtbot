import * as React from 'react';
import { useState } from 'react';
import styles from './MicrosoftChatBot.module.scss';
import type { IMicrosoftChatBotProps } from './IMicrosoftChatBotProps';
import { FaPaperPlane, FaRobot, FaTimes } from 'react-icons/fa';

interface IMessage {
  from: 'bot' | 'user';
  text: string;
  time: string;
}

const MicrosoftChatBot: React.FC<IMicrosoftChatBotProps> = ({ hasTeamsContext, userDisplayName }) => {
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
    }
  ]);
  const [isOpen, setIsOpen] = useState<boolean>(false);

  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setInputValue(event.target.value);
  };

  const handleSend = () => {
    if (!inputValue.trim()) return;

    const newMessage: IMessage = {
      from: 'user',
      text: inputValue.trim(),
      time: new Date().toLocaleString()
    };

    setMessages(prevMessages => [...prevMessages, newMessage]);
    setInputValue('');
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
      {/* Avatar icon */}
      {!isOpen && (
        <div className={styles.avatarIcon} onClick={toggleChatBot} aria-label="Open chat">
          <FaRobot size={32} />
        </div>
      )}

      {/* Chat window */}
      {isOpen && (
        <div className={`${styles.microsoftChatBot} ${hasTeamsContext ? styles.teams : ''}`}>
          <div className={styles.header}>
            <img
              src="https://your-company-logo-url.com/logo.png"
              alt="Company Logo"
              className={styles.logo}
            />
            <div>
            
              <h1>Welcome {userDisplayName}!</h1>
              <h3>RSM Support Center</h3>
            </div>

            <div className={styles.closeIcon} onClick={toggleChatBot} aria-label="Close chat">
              <FaTimes size={20} />
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

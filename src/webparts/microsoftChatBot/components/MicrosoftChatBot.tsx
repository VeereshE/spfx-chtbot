import * as React from "react";
import { useState } from "react";
import styles from "./MicrosoftChatBot.module.scss";
import type { IMicrosoftChatBotProps } from "./IMicrosoftChatBotProps";
import { FaPaperPlane, FaRobot, FaTimes } from "react-icons/fa";
import { MSGraphClientV3 } from "@microsoft/sp-http";

interface IMessage {
  from: "bot" | "user";
  text: string;
  time: string;
}

const MicrosoftChatBot: React.FC<IMicrosoftChatBotProps> = ({
  hasTeamsContext,
  userDisplayName,
  context,
}) => {
  const [inputValue, setInputValue] = useState<string>("");
  const [messages, setMessages] = useState<IMessage[]>([]);
  const [isOpen, setIsOpen] = useState<boolean>(false);
  const [chatId, setChatId] = useState<string | null>(null);
  const [userSuggestions, setUserSuggestions] = useState<any[]>([]);
  const [showSuggestions, setShowSuggestions] = useState<boolean>(false);
  const [selectedUser, setSelectedUser] = useState<any | null>(null);

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type

  const handleInputChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  ) => {
    const value = event.target.value;
    setInputValue(value);
     const atIndex = value.lastIndexOf("@");
    const searchTerm = value.slice(atIndex + 1).trim();

    if (atIndex !== -1 && (searchTerm.length === 0 || /^[a-zA-Z\s]*$/.test(searchTerm))) {
      try {
        const graphClient: MSGraphClientV3 =
          await context.msGraphClientFactory.getClient("3");
        const usersResponse = await graphClient
          .api("/users?$select=displayName,mail")
          .get();
        setUserSuggestions(usersResponse.value);
        setShowSuggestions(true);
      } catch (error) {
        console.error("Error fetching users:", error);
      }
    } else {
      setShowSuggestions(false);
    }
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const handleUserSelect = (user: any) => {
    const updatedText = inputValue.replace(/@$/, `@${user.displayName} `);
    setInputValue(updatedText);
    setSelectedUser(user);
    setShowSuggestions(false);
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const handleSend = async () => {
    if (!inputValue.trim()) return;

    const newMessage: IMessage = {
      from: "user",
      text: inputValue.trim(),
      time: new Date().toLocaleString(),
    };

    setMessages((prevMessages) => [...prevMessages, newMessage]);
    setInputValue("");

    try {
      const graphClient: MSGraphClientV3 =
        await context.msGraphClientFactory.getClient("3");

      // const recipientEmail = 'jake_madren@8dkwrz.onmicrosoft.com';
      let currentChatId = chatId;
      if (!currentChatId) {
        if (!selectedUser) {
          alert("Please mention a user using @ before sending a message.");
          setMessages([]);
          return;
        }

        const chatResponse = await graphClient.api("/chats").post({
          chatType: "oneOnOne",
          members: [
            {
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              roles: ["owner"],
              "user@odata.bind": "https://graph.microsoft.com/v1.0/me/mail",
            },
            {
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              roles: ["owner"],
              "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${selectedUser?.mail}')`,
            },
          ],
        });

        currentChatId = chatResponse.id;
        console.log(currentChatId);
        setChatId(currentChatId);
      }

      await graphClient.api(`/chats/${currentChatId}/messages`).post({
        body: {
          content: newMessage.text,
        },
      });

      console.log("Message Added ..............");
    } catch (error) {
      console.error("Error sending message via Graph API:", error);
    }
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const fetchUsers = async () => {
    try {
      const graphClient: MSGraphClientV3 =
        await context.msGraphClientFactory.getClient("3");
      const usersResponse = await graphClient.api("/users").get();
      console.log("Fetched Users:", usersResponse);
     // alert(usersResponse);
    } catch (error) {
      console.error("Error fetching users:", error);
    }
  };

  React.useEffect(() => {
    if (isOpen) {
      // eslint-disable-next-line no-void
      void fetchUsers();
    }
  }, [isOpen]);

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const handleKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
    if (event.key === "Enter") {
      // eslint-disable-next-line no-void
      void handleSend();
    }
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const toggleChatBot = () => {
    setIsOpen((prev) => !prev);
  };

  return (
    <section className={styles.chatWrapper}>
      {!isOpen && (
        <div
          className={styles.avatarIcon}
          onClick={toggleChatBot}
          aria-label="Open chat"
        >
          <FaRobot size={32} />
        </div>
      )}

      {isOpen && (
        <div
          className={`${styles.microsoftChatBot} ${
            hasTeamsContext ? styles.teams : ""
          }`}
        >
          <div className={styles.header}>
            <img
              src="https://rsmus.com/content/experience-fragments/rsm/us/en/site/header/master/_jcr_content/root/globalheader/mainnav/logo.coreimg.png/1648142668633/logo.png"
              alt="Company Logo"
              className={styles.logo}
            />
            <div>
              <h1>Welcome {userDisplayName}!</h1>
              <h3>RSM Support Center</h3>
            </div>
            <div
              className={styles.closeIcon}
              onClick={toggleChatBot}
              aria-label="Close chat"
            >
              <FaTimes size={32} />
            </div>
          </div>

          <div className={styles.chatContainer}>
            {messages.map((msg, index) => (
              <div
                key={index}
                className={`${styles.chatMessage} ${
                  msg.from === "bot" ? styles.botMessage : styles.userMessage
                }`}
              >
                <div className={styles.messageHeader}>
                  <span className={styles.userName}>
                    {msg.from === "bot" ? "Bot" : userDisplayName}
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
            {showSuggestions && userSuggestions.length > 0 && (
              <ul className={styles.suggestionsList}>
                {userSuggestions.map((user, index) => (
                  <li key={index} onClick={() => handleUserSelect(user)}>
                    {user.displayName}
                  </li>
                ))}
              </ul>
            )}
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

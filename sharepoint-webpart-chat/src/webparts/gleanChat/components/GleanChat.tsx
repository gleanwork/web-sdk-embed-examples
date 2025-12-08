import React, { useRef } from 'react';
import type { IGleanChatProps } from './IGleanChatProps';
import styles from './GleanChat.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';

const GleanChat = (props: IGleanChatProps) => {

  const containerRef = useRef<HTMLDivElement>(null);
  
  console.log(props.applicationId)

  console.log('Glean Chat Props:', props);
  // if you have a vanity url replace app.glean.com with your vanity url 
  SPComponentLoader.loadScript('https://app.glean.com/embedded-search-latest.min.js', {
    globalExportsName: 'EmbeddedSearch'
  }).then((EmbeddedSearch: any): void => {
    console.log('Library loaded');
    // Now you can use the library
    console.log(window.EmbeddedSearch); 
    window.EmbeddedSearch.renderChat(containerRef.current, {
      // customer backend url for Glean APIs
      // Replace this with your own backend URL
      backend: props.backendURL,
      // if you have set up a token service here is where you can pass the token
      //authToken: gleanToken,
      applicationId: props.applicationId || '',
      //onAuthTokenRequired: fetchToken,
      enableActivityLogging: true,
      themeVariant: 'dark'
    })
  }).catch((error) => {
    console.error('Error loading the library', error);
  });

  return (
    <div className={styles.gleanChat}>
      <div className="inner" ref={containerRef} />
    </div>
  );
};

export default GleanChat;
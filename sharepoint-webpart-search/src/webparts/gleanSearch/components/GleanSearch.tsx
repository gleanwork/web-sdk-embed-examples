import React, { useEffect } from 'react';
import type { IGleanSearchProps } from './IGleanSearchProps';
import styles from './GleanSearch.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';

const GleanSearch = (props: IGleanSearchProps) => {

  

  useEffect(() => {
    let attempts = 0; // Initialize attempt counter
    // Attempt to find the search box after the component mounts
    console.log('GleanSearch effect initializing');

    let intervalId: number | undefined = undefined; 

    // Function to find the search box
    const findSearchBox = () => {
      attempts++; // Increment the attempt counter
      console.log(`Attempt ${attempts} to find the search box`);

      // Find the search box
      const searchBox = document.querySelector('input[placeholder*="Search"]');

      // If the search box is found. attach the Glean search to it
      if (searchBox) {
        console.log('Found the search box:', searchBox);
        clearInterval(intervalId); // Stop trying to find the search box

        // Load the external library and attach it to the found search box
        SPComponentLoader.loadScript('https://app.glean.com/embedded-search-latest.min.js', {
          globalExportsName: 'EmbeddedSearch'
        }).then((EmbeddedSearch: any) => {
          console.log('EmbeddedSearch Library loaded');
          console.log("talking to backend " + props.backendURL);
          EmbeddedSearch.attach(searchBox, {
            // customer backend url for Glean APIs
            // Replace this with your own backend URL
            backend: props.backendURL,
            enableActivityLogging: true,
            // Add additional configurations as needed
          });
        }).catch((error) => {
          console.error('Error loading the library', error);
        });
      } else if (attempts > 10) {
        // Stop trying to find the search box after 10 attempts
        console.warn('Search box not found after 10 attempts');
        clearInterval(intervalId);
      } else {
        console.warn('Search box not found');
      }
    };

    // Set up an interval to try and find the search box every 500 milliseconds
    intervalId = setInterval(findSearchBox, 500);

    // Clean up the interval on component unmount
    return () => clearInterval(intervalId);
  }, []);

  return (
    <div className={styles.gleanSearch}>
      <div className="inner" />
    </div>
  );
};

export default GleanSearch;

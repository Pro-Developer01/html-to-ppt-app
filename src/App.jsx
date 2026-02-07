import { useState, useEffect } from 'react';
import './App.css';

function App() {
  const [htmlCode, setHtmlCode] = useState('');
  const [status, setStatus] = useState('Loading library...');
  const [isConverting, setIsConverting] = useState(false);
  const [libraryLoaded, setLibraryLoaded] = useState(false);

  useEffect(() => {
    // Load dom-to-pptx library dynamically from public folder
    const script = document.createElement('script');
    script.src = '/dom-to-pptx.bundle.js';
    script.async = true;
    script.onload = () => {
      setLibraryLoaded(true);
      setStatus('✅ Ready to convert...');
    };
    script.onerror = () => {
      setStatus('❌ Failed to load conversion library. Please refresh the page.');
    };
    document.body.appendChild(script);

    return () => {
      if (document.body.contains(script)) {
        document.body.removeChild(script);
      }
    };
  }, []);

  const convertToSlide = async () => {
    if (!htmlCode.trim()) {
      setStatus('❌ Please paste HTML code first');
      return;
    }

    setIsConverting(true);
    setStatus('Converting slide...');

    try {
      // Check if dom-to-pptx is available
      if (!window.domToPptx?.exportToPptx) {
        throw new Error('dom-to-pptx library not loaded');
      }

      const exportToPptx = window.domToPptx.exportToPptx;

      // Create temporary container
      const container = document.createElement('div');
      container.innerHTML = htmlCode;
      container.style.position = 'absolute';
      container.style.left = '-9999px';
      document.body.appendChild(container);

      setStatus('HTML loaded, rendering...');

      // Wait for fonts and styles to load
      await new Promise(resolve => setTimeout(resolve, 1000));

      // Find slide container (or use the whole container if no .slide-container)
      const slideElement = container.querySelector('.slide-container') || container.firstElementChild || container;

      if (!slideElement) {
        throw new Error('No valid HTML element found');
      }

      setStatus('Converting to PowerPoint...');

      // Convert to PPTX
      await exportToPptx(slideElement, {
        fileName: 'Slide.pptx',
        slideWidth: 10,
        slideHeight: 5.625
      });

      setStatus('✅ Success! Your PowerPoint file has been downloaded.');

      // Cleanup
      document.body.removeChild(container);

    } catch (error) {
      setStatus(`❌ Error: ${error.message}`);
      console.error(error);
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <div className="app">
      <div className="container">
        <h1>HTML to PowerPoint Converter</h1>
        <p className="subtitle">Paste your HTML code and convert to PPTX</p>
        
        <div className="info">
          <strong>📋 Instructions:</strong><br />
          1. Paste your HTML code in the text area below<br />
          2. Click "Convert to PowerPoint" button<br />
          3. Your PPTX file will be downloaded automatically
        </div>

        <textarea
          className="html-input"
          placeholder="Paste your HTML code here..."
          value={htmlCode}
          onChange={(e) => setHtmlCode(e.target.value)}
          rows={15}
        />

        <button 
          onClick={convertToSlide}
          disabled={isConverting || !libraryLoaded}
          className="convert-button"
        >
          {isConverting ? 'Converting...' : 'Convert to PowerPoint'}
        </button>
        
        <div className="status" dangerouslySetInnerHTML={{ __html: status }} />
      </div>
    </div>
  );
}

export default App;

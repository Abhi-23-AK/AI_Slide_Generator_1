import React from "react";
import { SlideTheme } from "../utils/themeUtils";
import { Slide } from "../types";
import { EditableText } from "../layouts/EditableText";

interface SlideHeadingProps {
  title: string;
  icon_emoji?: string; // e.g. "🤖", "📊", "🚀"
  theme: SlideTheme;
  size?: 'large' | 'medium' | 'small';
  slide?: Slide;
  zoneKey?: string;
  color?: string;
  style?: React.CSSProperties;
  textStyle?: React.CSSProperties;
  emojiStyle?: React.CSSProperties;
}

const SlideHeading = ({ 
  title, 
  icon_emoji, 
  theme,
  size = 'large',
  slide,
  zoneKey,
  color,
  style,
  textStyle,
  emojiStyle,
}: SlideHeadingProps) => {
  
  const fontSizes = {
    large: '36px',   // main slide title
    medium: '24px',  // section heading
    small: '18px'    // card heading
  };

  const emojiSizes = {
    large: '42px', // ~1.2x of 36px
    medium: '28px', // ~1.2x of 24px
    small: '22px'   // ~1.2x of 18px
  };

  // Dynamic font size shrinking for slide titles to keep them on a single line
  let currentFontSize = textStyle?.fontSize || fontSizes[size];
  if (!textStyle?.fontSize) {
    const len = title.length;
    if (size === 'large') {
      if (len > 45) {
        currentFontSize = '20px';
      } else if (len > 35) {
        currentFontSize = '24px';
      } else if (len > 25) {
        currentFontSize = '28px';
      }
    } else if (size === 'medium') {
      if (len > 45) {
        currentFontSize = '16px';
      } else if (len > 35) {
        currentFontSize = '18px';
      } else if (len > 25) {
        currentFontSize = '20px';
      }
    }
  }
  
  // Calculate dynamic emoji size if custom fontSize is provided
  let currentEmojiSize = emojiSizes[size];
  if (currentFontSize && typeof currentFontSize === 'string') {
    const fs = parseFloat(currentFontSize);
    if (!isNaN(fs)) {
      currentEmojiSize = `${Math.round(fs * 1.2)}px`;
    }
  }

  const headingStyle: React.CSSProperties = {
    fontSize: currentFontSize,
    fontWeight: theme.headingWeight as any,
    fontFamily: theme.headingFont,
    color: color || theme.textHeading,
    textTransform: theme.headingTransform as any,
    lineHeight: 1.2,
    margin: 0,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    flex: 1,
    minWidth: 0,
    ...textStyle
  };

  return (
    <div style={{
      display: 'flex',
      alignItems: 'center',
      gap: '12px',
      marginBottom: '20px',
      width: '100%',
      ...style
    }}>
      {/* EMOJI ICON */}
      {icon_emoji && (
        <span style={{
          fontSize: currentEmojiSize,
          lineHeight: 1,
          flexShrink: 0,
          ...emojiStyle
        }}>
          {icon_emoji}
        </span>
      )}
      
      {/* HEADING TEXT */}
      {slide && zoneKey ? (
        <EditableText
          slide={slide}
          zoneKey={zoneKey}
          value={title}
          as="h1"
          style={headingStyle}
        />
      ) : (
        <h1 style={headingStyle}>
          {title}
        </h1>
      )}
    </div>
  );
};

export default SlideHeading;

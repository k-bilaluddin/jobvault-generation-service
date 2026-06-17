/**
 * Post-processes rendered DOCX XML to convert **markdown bold** markers
 * into proper Word bold runs (<w:b/>).
 *
 * docxtemplater substitutes {text} as a single <w:r> element. This function
 * finds any run whose <w:t> contains ** markers and splits it into alternating
 * normal / bold runs, preserving the original run properties (<w:rPr>).
 */
export function processBoldMarkdown(xml: string): string {
  return xml.replace(
    /<w:r>((?:<w:rPr>[\s\S]*?<\/w:rPr>)?)<w:t(?:[^>]*)>([\s\S]*?)<\/w:t><\/w:r>/g,
    (match, rPrBlock: string, text: string) => {
      if (!text.includes('**')) return match;

      const boldRPr = rPrBlock
        ? rPrBlock.includes('<w:b/>')
          ? rPrBlock
          : rPrBlock.replace('</w:rPr>', '<w:b/><w:bCs/></w:rPr>')
        : '<w:rPr><w:b/><w:bCs/></w:rPr>';

      return text
        .split('**')
        .map((part, i) => {
          if (!part) return '';
          const pr = i % 2 === 1 ? boldRPr : rPrBlock;
          return `<w:r>${pr}<w:t xml:space="preserve">${part}</w:t></w:r>`;
        })
        .join('');
    },
  );
}

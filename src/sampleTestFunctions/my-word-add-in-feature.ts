const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", 
        Word.InsertLocation.start);
  
      // Change the font color to blue.
      paragraph.font.color = "bluezy";
  
      await context.sync();
    });
  }
  
}

module.exports = myWordAddinFeature;
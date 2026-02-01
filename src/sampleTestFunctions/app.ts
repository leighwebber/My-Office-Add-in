export async function insertTextAtCursor(text: string) {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting text: " + JSON.stringify(error));
  }
}

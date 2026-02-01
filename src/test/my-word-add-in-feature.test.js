const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("/src/sampleTestFunctions/my-word-add-in-feature");
import { processMessage, openDialog, tryCatch, insertStageDiagram, insertIcon,
    insertParagraph, getSelectionText } from '/src/taskpane/taskpaneFunctions.ts'

// const insertParagraph = require('/src/taskpane/insertParagraph');

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },        
      },
      range: {
          load: function(text){
            return text: text,
          }
      },
      getSelection: function () {
          return {
            text: "This is the selected text"
          }
        },
      }
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
    start: "start"
  },
  // Mock the Word.run function.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;
/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("insert paragraph", async function () {
    await insertParagraph();
  })
  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("bluezy");
  })

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  })
  test("get selected text", async function (){
    await getSelectionText();
    expect(wordMock.context.document.getSelection.text).toBe("This is the selected text.");
  })
})

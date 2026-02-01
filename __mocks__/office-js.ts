// __mocks__/office-js.ts
export const Word = {
    run: jest.fn((callback) => {
        const context = {
            sync: jest.fn().mockResolvedValue(true),
            document: {
                body: {
                    insertParagraph: jest.fn()
                }
            }
        };
        return callback(context);
    })
};

// Also mock the Office global if necessary
(global as any).Office = {
    context: {}
};
(global as any).Word = Word;

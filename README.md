## Sport Works Product Description Translator and Formatter

This project is designed to help Sport Works, a family-owned sporting goods company, manage their product descriptions more efficiently. They get tons of new products every year, and translating and formatting all those descriptions for their website can be a real headache. Standard online translation tools often don't cut it and don't offer any formatting options. This app provides a simple GUI tool to make the whole process much smoother.

### What it does

The app uses an Excel spreadsheet as input. Here's how the spreadsheet should be set up:

*   **Row 1:** Product Codes
*   **Row 2:** All the English product info in an unformatted manner (descriptions, attributes, and compositions)

The app then translates and formats this English product information into Turkish using the ChatGPT API (model 4o). This provides a simple, automated, and low-cost solution to a time-consuming problem.

### How To Run
TODO

### Future Improvements
* Allow the user to pick the model they want to use.
* Allow more insight into prompting.
* Allow the user to pick the language they are translating from and to.
* Clean up UI.
* Add concurrency for large data processing.

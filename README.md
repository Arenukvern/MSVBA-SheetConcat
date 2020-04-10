# VBA-SheetConcat

The macro writes data from all sheets in a book, creates a new sheet and transfers all data to a new sheet. If the book is saved in .xls format then a new book is created and the data is transferred to the new book
## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

To install you need to download the file

```
MSVBA-SheetConcat
```

### Installing

<p>Look at the top of your Excel Window if you see the word “Developer” in the menu options, then you are ready to go.  You can skip straight ahead to the next step. However, if the Developer Ribbon is not there, just follow these instructions.</p>
<p>File -> Options -> Customize Ribbon</p>
A new window will open, ensure the Developer option is ticked in the box.
Click OK.

Find the Visual Basic Editor within the Developer Ribbon
Developer -> Visual Basic

Open the folder with downloaded files, select all files and drag and drop into a book

![alt text](https://dl4.joxi.net/drive/2020/04/10/0042/2948/2775940/40/1e30f89975.jpg "For exampel")


## Running the tests

There is way to run added Macro:

Developer -> Code -> Macros. Select Macro "SheetConcat", Click Run.

### Break down into end to end tests

If you have data on different sheets and you would like them put together to one sheet, you can use the Macro "SheetConcat".

Open the book with the sheets which need to merge and run the macro. After starting, the macro will create a new sheet and transfer to it all the data from all sheets sequentially one after another.


## Deployment

Add additional notes about how to deploy this on a live system

## Built With

* [Dropwizard](http://www.dropwizard.io/1.0.2/docs/) - The web framework used
* [Maven](https://maven.apache.org/) - Dependency Management
* [ROME](https://rometools.github.io/rome/) - Used to generate RSS Feeds

## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/PurpleBooth/b24679402957c63ec426) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags). 

## Authors

* **Anton Malofeev** - *Initial work* - [Arenukvern](https://github.com/Arenukvern)
* **Mihail Melnikov** - *Assist work* - [mixev](https://github.com/mixev)

See also the list of [contributors](https://github.com/your/project/contributors) who participated in this project.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* You can be the first to whom we will be grateful



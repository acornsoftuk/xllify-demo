# xllify-demo

This repo demonstrates using the [xllify-build](https://github.com/marketplace/actions/xllify-build) action to package some simple custom functions into an .xll. The workflow defined in [build.yaml](https://github.com/acornsoftuk/xllify-demo/blob/main/.github/workflows/build.yaml).

The action handles the build and packaging process. An .xll file for Microsoft Excel on Windows is attached when a release is created in this repo. See https://xllify.com for more info.

For more information about the action, see https://github.com/acornsoftuk/xllify-build or the workflow in this repo.

For comprehensive documentation of all available functions, see [DOCS.md](./DOCS.md). Note that xllifyAI wrote some of this code so best not to use these for anything other than an illustration, unless of course you know what you're doing.

## Testing

If you include a `tests/` directory in your repository, the xllify-build action will automatically discover and run your tests during the build process. If no `tests/` directory exists, the test runner will be skipped and the build will proceed normally. This makes testing completely optional but easy to add when you're ready.

## Installation

After downloading and installing from the [releases page](https://github.com/acornsoftuk/xllify-demo/releases/latest), Excel should show the shiny new functions.

![Insert function](./screenshots/insert.png)
![Function preview](./screenshots/preview.png)
![All](./screenshots/all.png)

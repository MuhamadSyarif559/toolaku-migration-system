# ToolakuMigrationSystem

This project was generated using [Angular CLI](https://github.com/angular/angular-cli) version 21.1.1.

## Development server

To start a local development server, run:

```bash
ng serve
```

Once the server is running, open your browser and navigate to `http://localhost:4200/`. The application will automatically reload whenever you modify any of the source files.

## Data Cleaning API

The Excel cleaning feature uses the local Node API in `server/index.js` and the Python scripts in `server/python/`.

1. Install Node dependencies:

```bash
npm install
```

2. Install Python dependencies:

```bash
python -m pip install -r requirements.txt
```

3. Start the local API:

```bash
npm run api:start
```

4. Start the Angular UI in another terminal if needed:

```bash
npm start
```

The API listens on `http://localhost:3001`. The frontend data-cleaning page calls:

- `GET /api/data-cleaning/catalog`
- `POST /api/data-cleaning/run`

To add another cleaner, add a new Python script under `server/python/` and register it in `DATA_CLEANING_TOOLS` inside `server/index.js`.

## Code scaffolding

Angular CLI includes powerful code scaffolding tools. To generate a new component, run:

```bash
ng generate component component-name
```

For a complete list of available schematics (such as `components`, `directives`, or `pipes`), run:

```bash
ng generate --help
```

## Building

To build the project run:

```bash
ng build
```

This will compile your project and store the build artifacts in the `dist/` directory. By default, the production build optimizes your application for performance and speed.

## Running unit tests

To execute unit tests with the [Vitest](https://vitest.dev/) test runner, use the following command:

```bash
ng test
```

## Running end-to-end tests

For end-to-end (e2e) testing, run:

```bash
ng e2e
```

Angular CLI does not come with an end-to-end testing framework by default. You can choose one that suits your needs.

## Additional Resources

For more information on using the Angular CLI, including detailed command references, visit the [Angular CLI Overview and Command Reference](https://angular.dev/tools/cli) page.

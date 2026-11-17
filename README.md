# Environment Variables Manager

A [Power Platform ToolBox](https://github.com/PowerPlatformToolBox) tool to list, search and compare Dataverse environment variables across environments.

## Features

- **List every environment variable** in the connected environment, with its type, default value and current value.
- **Search** across display name, schema name, description and values (secret references are never searched or shown in the grid).
- **Filter** by type, solution, and value state (has a current value / using the default / no value set).
- **Compare environments**: assign up to five connections and see each environment side by side. Each variable is marked **Same**, **Different** or **Missing**, and you can filter to only the differences. The comparison uses the *effective* value (the current value if set, otherwise the default). JSON values are compared structurally, and Boolean values case-insensitively.
- **Copy missing variables**: when comparing, a variable that's missing from an environment can be copied there from one that has it — hover the *Not in this environment* cell and click **Copy here**, use **Copy to …** in the details view, or **Copy missing** in the top bar to copy every missing variable shown. The schema name, display name, description, type and default value are copied (the current value only if you ask), into a solution you choose. Secret and Data source variables are skipped. Production asks for confirmation first.
- **Export**: **Export** saves the variables currently shown (search and filters applied) as Excel (.xlsx), CSV or JSON. In compare mode each environment gets its own current and default value columns. Secret variables export their Key Vault reference.
- **Details view**: click a row to see the full values for every environment, with JSON pretty-printed and copy buttons.
- **Create variables**: **Create variable** creates a Text, Number, Yes/No or JSON variable in a chosen environment and unmanaged solution, with the schema name prefix taken from the solution's publisher, and optional default and current values. Production asks for confirmation first.
- **Edit values**: hover a value in the grid and click its pencil to jump straight to its editor, or open the details view to edit the current or default value in any environment, remove a current value (so the default applies), or clear a default. Input is checked against the variable's type (number, Yes/No, valid JSON). Changes to a Production environment ask for confirmation first, and editing the default of a variable from a managed solution shows a warning. Secret values can't be edited here.

## Usage

1. Open the tool in PPTB and pick a connection.
2. To compare, add more connections from the connection footer (slots 2–5). The grid reloads automatically.
3. Use the search box and filters to narrow the list, and click a row for full details.

Requires PPTB **1.2.7** or later (connection slots). The signed-in user needs read access to the `environmentvariabledefinition`, `environmentvariablevalue`, `solution` and `solutioncomponent` tables, and write access to `environmentvariablevalue` (and `environmentvariabledefinition` for defaults and new variables) to edit or create them. Adding a new variable to a solution also needs permission to customize that solution.

## What this tool changes

This tool only writes to Dataverse when you choose to, and never when it opens. In the environment you pick it can:

- **Set a current value**: updates the `environmentvariablevalue` record, or creates one if there isn't one yet.
- **Remove a current value**: deletes the `environmentvariablevalue` record, so the default value applies again.
- **Set or clear a default value**: updates `defaultvalue` on the `environmentvariabledefinition`.
- **Create a variable**: creates an `environmentvariabledefinition` (and an optional current value) and can add it to an unmanaged solution with `AddSolutionComponent`.

It never deletes variable definitions or edits Secret or Data source values. In Production environments every change asks for confirmation first. Changes can't be undone from the tool. Use **Export** to save the current values first if you may need to restore them.

Exports (Excel, CSV, JSON) are written only to a file you choose.

## Development

```bash
npm install
npm run dev-watch   # dev build with sourcemaps, rebuilds on change
npm run build       # production build
npm run validate    # pptb-validate
```

To debug, open PPTB Settings and turn on **Show Debug Menu**. Then go to **Debug** → **Load Local Tool** and select this project's root folder. There's no hot reload, so close and reopen the tool tab after each rebuild.

## License

MIT

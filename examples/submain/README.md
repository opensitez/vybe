# Sub Main Example

This example demonstrates a VB project that starts with `Sub Main` instead of a form.

## Key Features

- **StartupObject**: Set to "Sub Main" in the project file
- **Console Application**: Runs code without showing any forms
- **Module-based**: Code is in Module1.vb rather than a form

## Project Structure

- `Module1.vb` - Contains the Main() subroutine that executes when the project starts
- `SubMain.vbproj` - Project file with `<StartupObject>Sub Main</StartupObject>`

## Testing New Functions

This example also tests the newly implemented VB functions:
- `CCur()` - Currency conversion with 4 decimal places
- `CVar()` - Variant conversion
- `IsNull()` - Check for NULL values
- `App` object - Application properties (Path, Title, etc.)
- `Screen` object - Screen/display properties (Width, Height, etc.)

## Running

Open the project in the vybe editor and click Run. The output will appear in the console.

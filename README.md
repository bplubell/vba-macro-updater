# Motivation

The intent of this macro is to update other macros that are in-session (currently being edited in the embedded VBA editor). It remove all modules, classes, and forms and replace them with files from disk. This separates the source code from the code that is running in the application which has some great advantages.

- Prevent data loss in case of macro corruption (such as when running an infinite loop)
- Reuse code - modules, classes, and forms in the `Shared` folder are automatically added to all in-session macros
- Track in version control software (such as git) for confident experimentation and rich history
- Enable use of better code editors (such as VSCode or Atom)

# Supported applications

At the moment, this is only tested with SOLIDWORKS. However, it should be extendable to the Microsoft Office suite.

# Getting started with SOLIDWORKS

## The sample

1. Download the "sample" folder
2. Create a new macro in the host application (such as SOLIDWORKS)
    - The location of the macro does not matter
    - The filename does not matter
3. Rename the project to `MessagingMacroProject`
    - Must match the folder in the sample folder with `Project` appended
    - ![Renamed macro project](/docs/images/sample_3_renamed-project.png)
4. Open the `MacroUpdater.swp` macro in the sample folder
5. Execute the `main` subroutine in the MacroUpdater project
6. Note that the `MessagingMacroProject` modules have been updated

## An existing macro

1. Create a new folder named after the project
2. Export the classes, modules, and forms from the editor to the new folder
    - Export available in the right-click menu:
    - ![Export command](/docs/images/existing_2_export.png)
3. Copy the `MacroUpdater.swp` to be a sibling of the project folder
4. [Optional] Created a `Shared` folder for code to share between more projects
    - Move/export any classes, modules, or forms that should be shared

After setting up, the folder structure should look like the following.

![Final directory structure](/docs/images/setup_directory-structure.png)

# Using with SOLIDWORKS

Once setup, edit the code in the project folder (in any text editor). Update the code in the application by executing the `Main` subroutine in the MacroUpdater project.

## Naming convention

There is a necessary naming convention that enables the MacroUpdater to "just work":

- There must exist on disk, as a sibling to "MacroUpdater.swp", a folder with the *project* name
- The name of the project (in the VBA editor) must match the name of the folder on disk with "Project" appended
- The name of the macro file does NOT matter but is typically the same name as the project folder on disk

# Contributing

Contributions are welcome, especially regarding Office support. Fork the repo, create a feature branch, make good edits, and submit a pull request!

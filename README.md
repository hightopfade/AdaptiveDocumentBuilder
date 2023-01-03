# Setup
Ensure that you have "Trust access to the VBA project object model" checked under
Trust Center -> Macro Settings

# Usage
To use module, first `import-module .\adb.psm1` this will import two functions
  1. ListAdversaries
  2. CreateDocument

## Commands
`ListAdversaries` will list the adversaries that are available to emulate
`CreateDocument` will create the document(s) based on the adversary chosen

## ListAdversary Example
```
PS C:\Users\ebfe\Desktop\adb> ListAdversaries

Name
----
test1
test2
test3
```

## CreateDocument Examples
The `CreateDocument` function takes 3 parameters
  1. `c` - this is the requested number of documents to be created (**NOT** required)
  2. `adversary` - this is the requested adversary to emulate (**REQUIRED**)
  3. `out` - this is the output directory to drop all of the requested files. default operation is to drop the files into `$PSScriptRoot`. if `out` is specified and the directory does not exist, it will be created (**NOT** required)
  4. `pass` - password you'd like to set on the document (**NOT** required)

```
PS C:\Users\ebfe\Desktop\adb> CreateDocument -c 1 -adversary test1

[*] Creating 1 documents...
```

```
PS C:\Users\ebfe\Desktop\adb> CreateDocument -c 5 -adversary test12

[*] Creating 5 documents...
```

```
PS C:\Users\ebfe\Desktop\adb> CreateDocument -c 5 -adversary test3 -out C:\Users\ebfe\Desktop\out

[*] Creating 5 documents...
[*] Specified output directory does NOT exist
[+] Creating 'C:\Users\ebfe\Desktop\out' now...
```

```
PS C:\Users\ebfe\Desktop\adb> CreateDocument -c 1 -adversary test1 -pass 'password'

[*] Creating '1' document(s)...
[!] Log does not exist
[+] Creating log file now...
[*] Created 'C:\Users\ebfe\Desktop\adb\Your_Invoice_4331.docm'...
```

# Development Notes
Ensure that the adversary directory is populated with a `Doc1.docx` or `Book1.xlsx`. This file is used as the template during the document creation process.

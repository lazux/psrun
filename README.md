# psrun

*Runs a PowerShell-Script while suppressing the implied console window*

PowerShell scripts are typically executed in the PowerShell console, resulting in the appearance of a console window. Even when using the `-WindowStyle Hidden` parameter to hide the console, a brief window flash still occurs. The purpose of this script is to completely suppress the console window and introduce additional useful features.

## There are many such scripts, so why have another one?

While there are numerous scripts available, the motivation for this "launcher" lies in avoiding the task of hiding the console window within the script itself. Additionally, a preference exists against the integration of mixed code (PowerShell / C#) and the utilization of "Scheduled Tasks", which can be cumbersome. This launcher aims to address the console window issue universally without requiring extensive configuration.

## So what are my requirements for a Powershell "launcher"?

It should be designed to meet the following requirements:

- Usability with dedicated board resources or widely used tools

- Capability to hide/show the console window and run target scripts with elevated privileges

- Acceptance of a variable number of parameters, forwarding them to the target script according to type

## How have these requirements been implemented?

The implementation utilizes VBScript, leveraging the underlying WSH available on all Windows systems following these principles:

- The path to the PowerShell script is specified via the `/File` parameter, supporting both relative paths and absolute paths, e.g. `/File "foo\bar.ps1"`

- The `/Hide` switch reliably suppresses the PowerShell console window

- The `/Elevate` switch ensures execution with elevated permissions

- Various parameter types intended for the target script are accepted and passed on:
  **Flags**: Value-less switches, e.g. `/foo`
  **Parameters**: Switchless, positional values, e.g. `"bar"`
  **Options**: Switch-bound values, e.g. `/foo:"bar"`

- Type-appropriate formatting of parameters is supported:
  **Strings** are enclosed in quotation marks
  **Integers** and **decimal numbers** are not enclosed in quotation marks
  **Multiple options** with switches of the same name e.g. `/foo:"bar" /foo:5 /foo:"baz"` are passed as a PowerShell-valid array `-foo "bar",5,"baz"`

> [!NOTE]  
> Decimal numbers must follow the en/us notation (with a dot as the decimal separator), e.g. 1234.56. If a different decimal separator is necessary, it should be entered as a string in quotation marks e.g., "1.234,56", as commas outside of strings are interpreted as value delimiters by the command line.

## Examples

- `wscript psrun.vbs /Command:"calc.exe" /Hide`

- `wscript psrun.vbs /NoExit /Command:"Get-ChildItem C:\Windows"`

- `wscript psrun.vbs /File:"D:\foo\bar.ps1" /Elevate`

- `wscript psrun.vbs /File:"bar.ps1" /Voo:2.5 /Voo:"baz" /Voo:7 /Doo`

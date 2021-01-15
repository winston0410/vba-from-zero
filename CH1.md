# Type and variable

## Type in VBA

<https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary>

## Declare variable in VBA

<https://bettersolutions.com/vba/variables/lifetime-scope.htm>

### Create variable with `Dim`

- If type is not set, variable is default to `Variant`

- Variables have three possible scopes, **procedure-level**, **module-level** and **global level**

**procedure-level**

Only Accessible **in the same procedure**.

```vba
Public Sub ProcedureOne()
Dim data As Integer
End Sub
```

**module-level**

The variable is declared in global without being set as `public`. Accessible **in the same module**.

```vba
Dim data As Integer
/* Is equal to */
Private data As Integer
```

**global level**

Accessible **across modules**.

```vba
Public data As String
```

### Declaring multiple variable with types

```vba
Dim sFirstName As String, sLastName As String
```

This will make the first variable to be `Variant`

```vba
Dim sFirstName, sLastName As String
```

### Abbreviation

```
Dim firstName$
```

is equal to

```
Dim firstName As String
```

<https://bettersolutions.com/vba/variables/abbreviations.htm>

## Initialising variables

Without initialising variables, default will be used.

<https://bettersolutions.com/vba/variables/initialising.htm>

## Mutating variable

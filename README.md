# VBA-LocalStorage

_Note: Under development_

Store key-value information for Workbooks/Documents locally with moderate security.
This can be useful for cases such as user-specific preferences, application state, and, in trusted environments, authentication information.
Stored values are encrypted with either an `ApplicationKey`, which is stored in VBA with the workbook, or a user-supplied password.
With the default `ApplicationKey` approach, the values are secure while at-rest when separate from the workbook, but if someone gains access to the computer and workbook is unprotected, they will be able to access the stored values.
This is a reasonable baseline for security and generally matches the approach used with browsers.
If additional security is desired, the user can elect to use a user-supplied password for encryption which would secure the values at-rest and would require entering the password to decrypt the values even with access to the computer and workbook.

## Overview

1. __Request Storage__ On first access, a dialog is displayed requesting usage of local storage for the given workbook/document. If the user rejects local storage, this preference is saved and in-memory storage will be used. If the user accepts, encryption with `ApplicationKey` will be used. Finally, the user has an option to accept with user-supplied password
2. On workbook/document re-open, if encrypted with user-supplied password, display __Unlock Storage__ user form, otherwise decrypts automatically

## Usage

Follows browser's `localStorage` approach with `GetItem`, `SetItem`, `RemoveItem`, and `Clear`

```vb
Private Token As String

Sub Login()
  Token = LocalStorage.GetItem("token")

  If Token = "" Then
    Token = "..."
    LocalStorage.SetItem "token", Token
  End If
End Sub

Sub Logout()
  Token = ""
  LocalStorage.RemoveItem "token"
End Sub

Sub Cleanup()
  LocalStorage.Clear
End Sub
```

## Installation

1. Import `LocalStorage.bas`, `RequestStorage.frm`, and `UnlockStorage.frm`
2. Set unique and strong value for `ApplicationKey` constant in `LocalStorage.bas`

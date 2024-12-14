# VBScript: Silent Failure in Parameter Validation

This example demonstrates a common error in VBScript where improper parameter validation can lead to a silent failure rather than a clear error message. The script defines a function that attempts to handle various input types, but it doesn't effectively handle cases where an unexpected error occurs.

The `bug.vbs` file contains the problematic code. The `bugSolution.vbs` file offers a solution that enhances error handling using `On Error Resume Next` and explicit error checks to provide better feedback to the user.

## How to reproduce the bug

1. Run the `bug.vbs` script.
2. Call the `MyFunction` with a parameter that doesn't fall into any of the handled cases. This might involve passing a non-numeric string that cannot be implicitly converted, or passing an object.  No obvious error will be thrown.
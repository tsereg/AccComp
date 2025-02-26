# AccComp
Microsoft Access schema comparer tool written with Claude.ai

This tool has been created using Claude.ai Sonnet 3.7 with Extended thinking model from a simple prompt. The generated code has then been redacted to remove warnings and build errors, and then uploaded back to Claude. 

The build errors were caused by the glitch (?) in Microsoft.Office.Interop.Access.Dao API: the Index object property Fields does not return a Fields object (field collection object), but IUnknown. There is also a bug (?) in implementation where iterating the indexes may cause an infinite loop. This has been handled by manual intervention in the generated code.

It took abount an hour to create the working tool, not including the time that was spent detecting and fixing the ACE DAO library glitch described above (this was already resolved on another project).

The conversation is shared here: https://claude.ai/share/9974e6f3-ff88-4143-ab37-ef427b96294a

## Requirements

The project (as given) compiles for 64-bit. Required are .NET 8 and Microsofot Access Database Engine (Microsoft ACE). The later should be installed in the bitness of the project. Microsoft.Office.Interop.Access.Dao NuGet package is used.


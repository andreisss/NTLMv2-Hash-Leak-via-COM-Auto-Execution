# NTLMv2-Hash-Leak-via-COM-Auto-Execution

Native auto-execution: Leverage login-time paths Windows trusts by default (Startup folder, Run-registry key).

Built-in COM objects: No exotic payloads or deprecated file types needed - just Shell.Application, Scripting.FileSystemObject and MSXML2.XMLHTTP and more COM objects.

Automatic NTLM auth: When your script points at a UNC share, Windows immediately tries to authenticate with NTLMv2.

The NTLM hash using native COM components and built-in Windows auto-execution logic. It is silent, reliable, and does not rely on deprecated or patched filetypes.




https://github.com/user-attachments/assets/f8e956e5-8fa2-4eb6-8642-dd2cfef013ad


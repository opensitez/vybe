# Threads & Processes Demo

Demonstrates the `System.Threading` and `System.Diagnostics` namespace features available in vybe Basic.

## Features Shown

| Feature | Description |
|---|---|
| `Thread.Sleep(ms)` | Pause execution for a given number of milliseconds |
| `Stopwatch` | High-resolution timer for measuring elapsed time |
| `Process.Start(cmd, args)` | Launch an external process |
| `Debug.WriteLine(msg)` | Write diagnostic output |
| `Debug.Assert(cond, msg)` | Assert a condition is true |

## Running

Open this project in the vybe IDE and press Run, or from the command line:

```
cargo run -p vybe_cli -- examples/threads_processes/Module1.bas
```

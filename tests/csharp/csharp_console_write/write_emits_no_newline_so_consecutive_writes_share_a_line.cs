// vybe-test: csharp/csharp_console_write/write_emits_no_newline_so_consecutive_writes_share_a_line
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

// console_write
Console.Write("a"); Console.Write("b"); Console.WriteLine();

// vybe-test: csharp/csharp_volatile_thread_memory/volatile_do_while_read_once_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

class FlagBox {
    public volatile int Value = 1;
}
var box = new FlagBox();
int count = 0;
do {
    count += box.Value;
    box.Value = 0;
} while (box.Value > 0);
Console.WriteLine(count);

// vybe-test: csharp/csharp_volatile_thread_memory/volatile_read_while_loop_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

class FlagBox {
    public volatile int Value = 3;
}
var box = new FlagBox();
int count = 0;
while (box.Value > 0) {
    count++;
    box.Value--;
}
Console.WriteLine(count);

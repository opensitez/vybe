// vybe-test: csharp/csharp_volatile_thread_memory/volatile_loop_write_accumulator_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

class FlagBox {
    public volatile int Value = 0;
}
var box = new FlagBox();
for (int i = 1; i <= 4; i++) box.Value += i;
Console.WriteLine(box.Value);

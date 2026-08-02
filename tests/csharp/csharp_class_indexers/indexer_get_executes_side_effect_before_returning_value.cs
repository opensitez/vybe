// vybe-test: csharp/csharp_class_indexers/indexer_get_executes_side_effect_before_returning_value
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Logger {
    int hits = 0;
    public int this[int key] {
        get { hits++; return key; }
    }
}
var logger = new Logger();
__Check((logger[5]).ToString(), "5");
__Check((logger.hits).ToString(), "1");

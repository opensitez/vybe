// vybe-test: kotlin/kotlin_class_init_sequences/test_secondary_constructor_with_default_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Config {
            val mode: String
            constructor() { mode = "auto" }
            constructor(raw: String) : this() { mode = raw }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Config().mode).toString(), "auto")
            __check((Config("manual").mode).toString(), "manual")
        }

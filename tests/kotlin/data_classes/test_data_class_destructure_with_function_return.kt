// vybe-test: kotlin/data_classes/test_data_class_destructure_with_function_return
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Record(val code: Int, val weight: Int)

        fun split(): Record {
            return Record(4, 5)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (code, weight) = split()
            __check((code).toString(), "4")
            __check((weight).toString(), "5")
            __check((code + weight).toString(), "9")
        }

// vybe-test: kotlin/local_classes/test_local_class_with_init_block
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local {
                val v: Int
                init { v = 3 }
            }
            __check((Local().v).toString(), "3")
        }

// vybe-test: kotlin/local_classes/test_local_class_multiple_instances
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local {
                val x = 1
            }
            val a = Local()
            val b = Local()
            __check((a.x + b.x).toString(), "2")
        }

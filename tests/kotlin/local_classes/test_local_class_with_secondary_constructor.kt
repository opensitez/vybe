// vybe-test: kotlin/local_classes/test_local_class_with_secondary_constructor
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
                constructor(v: Int) { this.v = v }
            }
            __check((Local(5).v).toString(), "5")
        }

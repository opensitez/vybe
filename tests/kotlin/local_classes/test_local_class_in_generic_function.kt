// vybe-test: kotlin/local_classes/test_local_class_in_generic_function
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun <T> use(v: T): String {
            class Holder {
                fun text() = v.toString()
            }
            return Holder().text()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((use(7)).toString(), "7")
        }

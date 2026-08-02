// vybe-test: kotlin/classes/test_object_declaration
// origin: languages/kotlin/tests/kotlin/test_classes.rs

object Logger {
            fun log(msg: String) {
                __check(("LOG: " + msg).toString(), "LOG: Started")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Logger.log("Started")
        }

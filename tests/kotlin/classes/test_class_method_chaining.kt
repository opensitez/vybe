// vybe-test: kotlin/classes/test_class_method_chaining
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Builder(var msg: String) {
            fun append(s: String): Builder {
                msg += s
                return this
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Builder("A")
            b.append("B").append("C")
            __check((b.msg).toString(), "ABC")
        }

// vybe-test: kotlin/kotlin_receiver_extension_ops/test_receiver_extension_with_this_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_receiver_extension_ops.rs

fun Int.doublePlusOne(): Int = this + this + 1

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.doublePlusOne()).toString(), "7")
        }

// vybe-test: kotlin/named_arguments/test_named_arguments_in_secondary_constructor_call
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

class Box {
            val value: Int
            val tag: String
            constructor(value: Int, tag: String = "x") {
                this.value = value
                this.tag = tag
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = Box(tag = "z", value = 9)
            __check((x.value).toString(), "9")
            __check((x.tag).toString(), "z")
        }

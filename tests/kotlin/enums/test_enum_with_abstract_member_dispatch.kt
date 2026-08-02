// vybe-test: kotlin/enums/test_enum_with_abstract_member_dispatch
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Operation {
            ADD {
                override fun apply(a: Int, b: Int): Int = a + b
            },
            SUBTRACT {
                override fun apply(a: Int, b: Int): Int = a - b
            },
            MULTIPLY {
                override fun apply(a: Int, b: Int): Int = a * b
            }
abstract fun apply(a: Int, b: Int): Int
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Operation.ADD.apply(4, 2)).toString(), "6")
            __check((Operation.SUBTRACT.apply(7, 3)).toString(), "4")
            __check((Operation.MULTIPLY.apply(3, 5)).toString(), "15")
        }

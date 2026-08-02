// vybe-test: kotlin/type_aliases/test_typealias_sequence_operations_preserve_lazy_contract
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias IntSequence = Sequence<Int>

        fun firstSquares(limit: Int): IntSequence {
            return generateSequence(0) { value ->
                if (value + 2 <= limit) value + 2 else null
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = firstSquares(6)
            __check((values.take(3).joinToString(",")).toString(), "2,4,6")
            __check((firstSquares(6).sum()).toString(), "12")
        }

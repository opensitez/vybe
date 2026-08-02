// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_to_mutable
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = sequenceOf(1, 2, 3).toMutableList()
            list.add(4)
            __check((list.joinToString(",")).toString(), "1,2,3,4")
        }

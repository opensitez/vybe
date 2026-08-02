// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_transform_with_index
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf("a", "b", "c")
            val withIndex = seq.withIndex().map { "${'$'}{it.index}:${'$'}{it.value}" }
            __check((withIndex.toList().joinToString(",")).toString(), "0:a,1:b,2:c")
        }

// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_zip
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = sequenceOf(1, 2, 3)
            val b = sequenceOf("x", "y", "z")
            val zipped = a.zip(b)
            __check((zipped.toList().joinToString(",") { (n, s) -> "$n$s" }).toString(), "1x,2y,3z")
        }

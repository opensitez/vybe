// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_with_mutable_output_reuse
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = mutableMapOf<Int, String>()
            listOf("x", "yy", "zzz").associateByTo(out, { it.length }) { it }
            listOf("a", "bb").associateByTo(out, { it.length }) { it + "!" }
            __check((out[1]).toString(), "x")
            __check((out[2]).toString(), "bb!")
        }

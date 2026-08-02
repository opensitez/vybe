// vybe-test: kotlin/kotlin_associate_apis/test_associate_to_mutable_seeded
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = linkedMapOf<Int, String>()
            out[1] = "seed"
            listOf("alpha", "bee").associateByTo(out, { it.length }) { it }
            __check((out[5]).toString(), "bee")
            __check((out.size).toString(), "2")
        }

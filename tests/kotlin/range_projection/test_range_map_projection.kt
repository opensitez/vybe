// vybe-test: kotlin/range_projection/test_range_map_projection
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scaled = (1..4).map { it * 2 }
            val names = (1..3).map { "p" + it }
            __check((scaled.joinToString(",")).toString(), "2,4,6,8")
            __check((names.joinToString(",")).toString(), "p1,p2,p3")
        }

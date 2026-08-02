// vybe-test: kotlin/advanced_features/test_advanced_object_expression_chain
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

interface Flag { fun value(): String }
fun make() = object : Flag { override fun value(): String = "go" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((make().value()).toString(), "go") }

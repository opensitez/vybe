// vybe-test: kotlin/scope/test_top_level_and_local_scope_isolation
// origin: languages/kotlin/tests/kotlin/test_scope.rs

val label = "global"

        fun echo(value: String): String {
            return label + ":" + value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val label = "local"
            __check((echo("one")).toString(), "global:one")
            __check((label).toString(), "local")
        }

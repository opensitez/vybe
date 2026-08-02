// vybe-test: kotlin/data_classes/test_data_class_multiple_instances_in_map_lookup_by_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Route(val from: Int, val to: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val route = Route(1, 2)
            val lookup = mapOf(route to "ok")
            val probe = route.copy(to = 2)
            __check((lookup[probe]).toString(), "ok")
        }

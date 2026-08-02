// vybe-test: kotlin/type_cast_edges/test_nested_smart_cast_in_while
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun main() {
            var value: Any = "start"
            var out = 0
            while (value is String) {
                out = value.length
                value = 10
            }
            println(out)
        }


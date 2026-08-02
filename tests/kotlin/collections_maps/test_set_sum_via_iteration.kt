// vybe-test: kotlin/collections_maps/test_set_sum_via_iteration
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val ids = setOf(3, 4, 5)
            var total = 0
            for (id in ids) {
                total += id
            }
            println(total)
            println(ids.size == 3)
        }


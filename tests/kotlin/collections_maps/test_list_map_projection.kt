// vybe-test: kotlin/collections_maps/test_list_map_projection
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val nums = listOf(1, 2, 3)
            val doubled = nums.map { it * 2 }
            var total = 0
            for (v in doubled) {
                total += v
            }
            println(doubled[0] + doubled[1] + doubled[2])
            println(total)
        }


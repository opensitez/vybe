// vybe-test: kotlin/collections_maps/test_map_key_membership_across_nested_collections
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val registry = mapOf(
                "admin" to listOf("read", "write"),
                "guest" to listOf("read")
            )
            var canWrite = false
            val role = "admin"
            for ((user, permissions) in registry) {
                if (user == role) {
                    for (perm in permissions) {
                        if (perm == "write") {
                            canWrite = true
                        }
                    }
                }
            }
            println(canWrite)
            println(registry.containsKey("guest"))
        }


use crate::helpers::run_prints;

#[test]
fn test_tuple_pair_creation_and_members() {
    let out = run_prints(r#"
        fun main() {
            val pair = Pair("x", 4)
            println(pair.first)
            println(pair.second)
        }
    "#);
    assert_eq!(out, &["x", "4"]);
}

#[test]
fn test_tuple_infix_to_constructs_pair() {
    let out = run_prints(r#"
        fun main() {
            val pair = "k" to 9
            println(pair.first)
            println(pair.second)
        }
    "#);
    assert_eq!(out, &["k", "9"]);
}

#[test]
fn test_tuple_pair_component_functions() {
    let out = run_prints(r#"
        fun main() {
            val pair = Pair(8, "v")
            println(pair.component1())
            println(pair.component2())
        }
    "#);
    assert_eq!(out, &["8", "v"]);
}

#[test]
fn test_tuple_pair_string_representation() {
    let out = run_prints(r#"
        fun main() {
            val pair = Pair("a", "b")
            println(pair)
        }
    "#);
    assert_eq!(out, &["(a, b)"]);
}

#[test]
fn test_tuple_pair_equality_and_inequality() {
    let out = run_prints(r#"
        fun main() {
            println(Pair(1, 2) == Pair(1, 2))
            println(Pair(1, 2) == Pair(2, 1))
            println(Pair(1, 2) != Pair(2, 1))
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_tuple_pair_null_components() {
    let out = run_prints(r#"
        fun main() {
            val pair: Pair<String?, Int?> = Pair(null, null)
            println(pair.first == null)
            println(pair.second == null)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_tuple_pair_as_map_key_lookup_with_identity_by_value() {
    let out = run_prints(r#"
        fun main() {
            val key = Pair("id", 3)
            val map = mapOf(key to "found")
            println(map[Pair("id", 3)])
            println(map[Pair("id", 4)] ?: "missing")
        }
    "#);
    assert_eq!(out, &["found", "missing"]);
}

#[test]
fn test_tuple_pair_in_set_uniqueness() {
    let out = run_prints(r#"
        fun main() {
            val seen = setOf(Pair(1, 2), Pair(1, 2), Pair(2, 1))
            println(seen.size)
            println(seen.contains(Pair(2, 1)))
            println(seen.contains(Pair(3, 4)))
        }
    "#);
    assert_eq!(out, &["2", "true", "false"]);
}

#[test]
fn test_tuple_nested_pair_unwrap_access() {
    let out = run_prints(r#"
        fun main() {
            val nested = Pair(Pair(1, 2), Pair(3, 4))
            println(nested.first.first)
            println(nested.second.second)
        }
    "#);
    assert_eq!(out, &["1", "4"]);
}

#[test]
fn test_tuple_pair_from_function_return() {
    let out = run_prints(r#"
        fun main() {
            fun make(): Pair<Int, Int> {
                return Pair(7, 11)
            }
            val (left, right) = make()
            println(left + right)
        }
    "#);
    assert_eq!(out, &["18"]);
}

#[test]
fn test_tuple_pair_in_function_parameter() {
    let out = run_prints(r#"
        fun main() {
            fun sum(pair: Pair<Int, Int>): Int {
                return pair.first + pair.second
            }
            println(sum(Pair(4, 6)))
            println(sum(8 to 1))
        }
    "#);
    assert_eq!(out, &["10", "9"]);
}

#[test]
fn test_tuple_pair_in_when_expression() {
    let out = run_prints(r#"
        fun main() {
            val pair = Pair("blue", 3)
            val label = when (pair) {
                Pair("red", 1) -> "red-one"
                Pair("blue", 3) -> "blue-three"
                else -> "other"
            }
            println(label)
        }
    "#);
    assert_eq!(out, &["blue-three"]);
}

#[test]
fn test_tuple_pair_of_lists_projection_to_map_key_values() {
    let out = run_prints(r#"
        fun main() {
            val rows = listOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            val labels = rows.toMap()
            println(labels["b"])
            println(labels["x"] ?: -1)
            println(labels.size)
        }
    "#);
    assert_eq!(out, &["2", "-1", "3"]);
}

#[test]
fn test_tuple_pair_with_mutable_ref_elements() {
    let out = run_prints(r#"
        fun main() {
            val a = mutableListOf(1)
            val b = mutableListOf(2)
            val pair = Pair(a, b)
            pair.first.add(3)
            pair.second.add(4)
            println(a.size)
            println(b.size)
            println(pair.first[1] + pair.second[1])
        }
    "#);
    assert_eq!(out, &["2", "2", "7"]);
}

#[test]
fn test_tuple_pair_iteration_as_two_value_sequence() {
    let out = run_prints(r#"
        fun main() {
            val data = listOf(Pair(1, 10), Pair(2, 20), Pair(3, 30))
            var total = 0
            for (i in data) {
                total += i.first
                total += i.second
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["66"]);
}

#[test]
fn test_tuple_pair_zip_default_no_transform_is_pair() {
    let out = run_prints(r#"
        fun main() {
            val left = listOf(1, 2, 3)
            val right = listOf("a", "b", "c")
            val zipped = left.zip(right)
            println(zipped.size)
            println(zipped[0])
            println(zipped[1].first + zipped[1].second.length)
        }
    "#);
    assert_eq!(out, &["3", "(1, a)", "3"]);
}

#[test]
fn test_tuple_pair_zip_is_empty_when_lengths_mismatch() {
    let out = run_prints(r#"
        fun main() {
            val zipped = listOf(1, 2).zip(listOf("x", "y", "z"))
            println(zipped.size)
            println(zipped[1].first)
            println(zipped[1].second)
        }
    "#);
    assert_eq!(out, &["2", "2", "y"]);
}

#[test]
fn test_tuple_pair_unzip_roundtrip() {
    let out = run_prints(r#"
        fun main() {
            val source = listOf(Pair(1, "a"), Pair(2, "b"), Pair(3, "c"))
            val (nums, chars) = source.unzip()
            println(nums.joinToString(","))
            println(chars.joinToString(""))
        }
    "#);
    assert_eq!(out, &["1,2,3", "abc"]);
}

#[test]
fn test_tuple_pair_arrayof_roundtrip() {
    let out = run_prints(r#"
        fun main() {
            val source = arrayOf(Pair(9, 8), Pair(7, 6))
            val list = source.toList()
            println(list[0].first + list[1].second)
            println(list[1].toString())
        }
    "#);
    assert_eq!(out, &["15", "(7, 6)"]);
}

#[test]
fn test_tuple_triple_creation_and_components() {
    let out = run_prints(r#"
        fun main() {
            val triple = Triple(1, "x", true)
            println(triple.first)
            println(triple.second)
            println(triple.third)
        }
    "#);
    assert_eq!(out, &["1", "x", "true"]);
}

#[test]
fn test_tuple_triple_component_functions() {
    let out = run_prints(r#"
        fun main() {
            val triple = Triple("a", 2, 3)
            println(triple.component1())
            println(triple.component2())
            println(triple.component3())
        }
    "#);
    assert_eq!(out, &["a", "2", "3"]);
}

#[test]
fn test_tuple_triple_to_string() {
    let out = run_prints(r#"
        fun main() {
            val triple = Triple(1, 2, 3)
            println(triple)
        }
    "#);
    assert_eq!(out, &["(1, 2, 3)"]);
}

#[test]
fn test_tuple_triple_equality() {
    let out = run_prints(r#"
        fun main() {
            println(Triple(1, 2, 3) == Triple(1, 2, 3))
            println(Triple(1, 2, 3) == Triple(3, 2, 1))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_tuple_triple_as_map_key() {
    let out = run_prints(r#"
        fun main() {
            val map = mapOf(Triple("a", 1, true) to "ok")
            println(map[Triple("a", 1, true)])
            println(map[Triple("a", 1, false)] ?: "missing")
        }
    "#);
    assert_eq!(out, &["ok", "missing"]);
}

#[test]
fn test_tuple_nested_tuple_unpacking_via_indexing() {
    let out = run_prints(r#"
        fun main() {
            val nested = Pair(Triple(1, 2, 3), Triple(4, 5, 6))
            println(nested.first.second + nested.second.first)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_tuple_zip_pairs_from_primitive_ranges() {
    let out = run_prints(r#"
        fun main() {
            val zipped = (1..4).zip(10 downTo 7)
            println(zipped.size)
            println(zipped[0])
            println(zipped[3])
        }
    "#);
    assert_eq!(out, &["4", "(1, 10)", "(4, 7)"]);
}

#[test]
fn test_tuple_triple_function_with_return_position() {
    let out = run_prints(r#"
        fun main() {
            fun stats(): Triple<Int, Int, Int> {
                return Triple(2, 4, 6)
            }
            val score = stats()
            println(score.third / score.second)
            println(score.first + score.second + score.third)
        }
    "#);
    assert_eq!(out, &["1", "12"]);
}

#[test]
fn test_tuple_triple_in_collection_sorting_stability() {
    let out = run_prints(r#"
        fun main() {
            val points = listOf(
                Triple("a", 3, 2),
                Triple("b", 1, 9)
            )
            println(points.sortedBy { it.second }.joinToString("|") { it.first })
        }
    "#);
    assert_eq!(out, &["b|a"]);
}

#[test]
fn test_tuple_pair_and_triple_mixed_collection_projection() {
    let out = run_prints(r#"
        fun main() {
            val mixed = listOf(
                Pair("p", 1),
                Triple("t", 2, 3)
            )
            println(mixed[0])
            println((mixed[1] as Triple<String, Int, Int>).second)
        }
    "#);
    assert_eq!(out, &["(p, 1)", "2"]);
}

#[test]
fn test_tuple_pair_values_replace_in_mutable_collection() {
    let out = run_prints(r#"
        fun main() {
            val pairs = mutableListOf(Pair(1, 2), Pair(3, 4))
            pairs[1] = Pair(5, 6)
            println(pairs[0])
            println(pairs[1])
            println(pairs.size)
        }
    "#);
    assert_eq!(out, &["(1, 2)", "(5, 6)", "2"]);
}

#[test]
fn test_pair_destructuring_captures_expected_values() {
    let out = run_prints(r#"
        fun main() {
            val (left, right) = Pair(8, 13)
            println(left)
            println(right)
        }
    "#);
    assert_eq!(out, &["8", "13"]);
}

#[test]
fn test_triple_destructuring_skips_first_component() {
    let out = run_prints(r#"
        fun main() {
            val (_, mid, last) = Triple("a", 14, 15)
            println(mid)
            println(last)
        }
    "#);
    assert_eq!(out, &["14", "15"]);
}

#[test]
fn test_destructuring_works_in_for_loop_over_pairs() {
    let out = run_prints(r#"
        fun main() {
            val rows = listOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            var total = ""
            for ((label, value) in rows) {
                total += "$label$value-"
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["a1-b2-c3-"]);
}

#[test]
fn test_destructuring_works_with_triple_in_while_like_rewrite() {
    let out = run_prints(r#"
        fun main() {
            var index = 0
            var total = ""
            val values = listOf(Triple("a", 1, 10), Triple("b", 2, 20))
            while (index < values.size) {
                val (_, left, right) = values[index]
                total += "$left:$right;"
                index++
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["1:10;2:20;"]);
}

#[test]
fn test_destructuring_assignment_uses_last_writer() {
    let out = run_prints(r#"
        fun main() {
            var pair = Pair(1, 2)
            var (first, second) = pair
            first = 9
            pair = Pair(first, second + 1)
            println(pair)
        }
    "#);
    assert_eq!(out, &["(9, 3)"]);
}

#[test]
fn test_pair_projection_after_mutation_remains_value_copy() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableListOf(1, 2)
            val (left, right) = Pair(source[0], source[1])
            source[0] = 9
            println(left)
            println(right)
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_tuple_with_nullable_components_in_destructure() {
    let out = run_prints(r#"
        fun main() {
            val pair: Pair<String?, String?> = Pair(null, "x")
            val (left, right) = pair
            println(left == null)
            println(right.length)
        }
    "#);
    assert_eq!(out, &["true", "1"]);
}

#[test]
fn test_triple_used_as_pair_like_in_map_with_projection() {
    let out = run_prints(r#"
        fun main() {
            val points = listOf(
                Triple("a", 1, 10),
                Triple("b", 2, 20)
            ).associateBy { it.first }
            val first = "a"
            val values = points[first]!!
            println(first)
            println(values.second)
            println(values.third)
        }
    "#);
    assert_eq!(out, &["a", "1", "10"]);
}

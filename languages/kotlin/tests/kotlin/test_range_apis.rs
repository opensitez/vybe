use crate::helpers::run_prints;

#[test]
fn test_range_inclusive_contains_boundaries() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..5
            println(r.first)
            println(r.last)
            println(1 in r)
            println(5 in r)
            println(6 in r)
        }
    "#,
    );
    assert_eq!(out, &["1", "5", "true", "true", "false"]);
}

#[test]
fn test_until_excludes_end() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1 until 5
            println(r.first)
            println(r.last)
            println(5 in r)
            println(4 in r)
            println(r.isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["1", "4", "false", "true", "false"]);
}

#[test]
fn test_down_to_descending_range() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 5 downTo 1
            println(r.first)
            println(r.last)
            println(5 in r)
            println(1 in r)
            println(6 in r)
            println(r.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["5", "1", "true", "true", "false", "5,4,3,2,1"]);
}

#[test]
fn test_range_step_even_spacing() {
    let out = run_prints(
        r#"
        fun main() {
            val r = (1..10).step(3)
            println(r.first)
            println(r.last)
            println(r.step)
            println(r.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1", "10", "3", "1,4,7,10"]);
}

#[test]
fn test_reverse_range() {
    let out = run_prints(
        r#"
        fun main() {
            val r = (1..5).reversed()
            println(r.first)
            println(r.last)
            println(r.step)
            println(r.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["5", "1", "-1", "5,4,3,2,1"]);
}

#[test]
fn test_range_is_empty_false() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..1
            println(r.isEmpty())
            println(r.contains(1))
            println(r.last)
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "1"]);
}

#[test]
fn test_range_singleton_down_to() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1 downTo 1
            println(r.toList().joinToString(","))
            println(r.isEmpty())
            println(r.first)
            println(r.last)
        }
    "#,
    );
    assert_eq!(out, &["1", "false", "1", "1"]);
}

#[test]
fn test_long_range_progression() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 10L downTo 6L
            println(r.first)
            println(r.last)
            println((r.step).toString())
            println(r.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["10", "6", "-1", "10,9,8,7,6"]);
}

#[test]
fn test_char_range_contains() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 'a'..'d'
            println(r.first)
            println(r.last)
            println('c' in r)
            println('x' in r)
        }
    "#,
    );
    assert_eq!(out, &["a", "d", "true", "false"]);
}

#[test]
fn test_char_range_to_list() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 'x'..'z'
            println(r.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["x,y,z"]);
}

#[test]
fn test_char_range_step_two() {
    let out = run_prints(
        r#"
        fun main() {
            val r = ('a'..'f').step(2)
            println(r.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,c,e"]);
}

#[test]
fn test_int_range_count_and_last_element() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 2..9
            var count = 0
            var last = 0
            for (v in r) { count++; last = v }
            println(count)
            println(last)
        }
    "#,
    );
    assert_eq!(out, &["8", "9"]);
}

#[test]
fn test_int_range_empty_when_step_non_overlap() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..5 step 1
            val d = 5 downTo 1
            println((1..5 step 0).isEmpty())
            println((1..5).step(-1).isEmpty())
            println(r.toList().size)
            println(d.toList().size)
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "5", "5"]);
}

#[test]
fn test_int_range_sum_via_fold() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..4
            val total = r.fold(0) { acc, value -> acc + value }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_int_range_reduce_via_terminal() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 2..5
            val value = r.reduce { acc, value -> acc * value }
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["120"]);
}

#[test]
fn test_range_any_all_none() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..6
            println(r.any { it % 2 == 0 })
            println(r.all { it > 0 })
            println(r.none { it < 0 })
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_range_last_of_down_to_step() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 9 downTo 3 step 3
            println(r.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9,6,3"]);
}

#[test]
fn test_range_with_negative_step_to_list() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 6 downTo 1
            println(r.step)
            println(r.toList().joinToString(";"))
        }
    "#,
    );
    assert_eq!(out, &["-1", "6;5;4;3;2;1"]);
}

#[test]
fn test_range_projection_of_first_last() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..10
            println(r.first)
            println(r.last)
            println(r.count())
        }
    "#,
    );
    assert_eq!(out, &["1", "10", "10"]);
}

#[test]
fn test_range_contains_many_points() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 10 downTo 1
            println(10 in r)
            println(1 in r)
            println(0 in r)
            println(11 in r)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "false"]);
}

#[test]
fn test_long_range_map_to_ints() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1L..3L
            val sum = r.map { it.toInt() }.sum()
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_char_range_reversed_has_expected_order() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 'c' downTo 'a'
            println(r.toList().joinToString(","))
            println(r.first)
            println(r.last)
        }
    "#,
    );
    assert_eq!(out, &["c,b,a", "c", "a"]);
}

#[test]
fn test_range_coerce_at_least_within() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..5
            val value = r.coerceIn(3)
            println(value)
            val under = r.coerceIn(0)
            println(under)
            val over = r.coerceIn(7)
            println(over)
        }
    "#,
    );
    assert_eq!(out, &["3", "1", "5"]);
}

#[test]
fn test_range_contains_open_ended_comparison() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 2..2
            println(r.first)
            println(r.last)
            println(r.contains(2))
            println(r.contains(3))
        }
    "#,
    );
    assert_eq!(out, &["2", "2", "true", "false"]);
}

#[test]
fn test_range_of_indices_used_with_list() {
    let out = run_prints(
        r#"
        fun main() {
            val data = listOf(10, 20, 30, 40)
            println(data.indices.toList().joinToString(","))
            println(data.indices.last)
            println(data[ data.indices.last ])
        }
    "#,
    );
    assert_eq!(out, &["0,1,2,3", "3", "40"]);
}

#[test]
fn test_range_to_typed_progression_default_step() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 3 until 3
            println(r.isEmpty())
            println(r.toList().size)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_range_drop_take_like_slice() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..10
            println(r.drop(2).take(3).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,4,5"]);
}

#[test]
fn test_range_join_to_string() {
    let out = run_prints(
        r#"
        fun main() {
            val r = 1..5
            println(r.joinToString(".") )
        }
    "#,
    );
    assert_eq!(out, &["1.2.3.4.5"]);
}

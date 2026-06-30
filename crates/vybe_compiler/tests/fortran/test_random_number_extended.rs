//! Extended random_number and random_seed: distribution bounds, put/get, reproducibility.
//! Distinct from compile-only probes in `test_intrinsics_extended.rs`.

fortran_cases! {
    random_number_scalar_in_unit_interval => {
        "program t\nreal :: r\ncall random_number(r)\nprint *, merge(1, 0, r >= 0.0 .and. r < 1.0)\nend program t\n",
        ["1"]
    };

    random_number_array_all_in_interval => {
        "program t\nreal :: r(5)\ncall random_number(r)\nprint *, merge(1, 0, all(r >= 0.0 .and. r < 1.0))\nend program t\n",
        ["1"]
    };

    random_seed_put_compile_and_mark => {
        "program t\ninteger :: seed(1) = [42]\ncall random_seed(put=seed)\nprint *, 1\nend program t\n",
        ["1"]
    };

    random_seed_get_returns_size => {
        "program t\ninteger :: n\ncall random_seed(size=n)\nprint *, merge(1, 0, n >= 1)\nend program t\n",
        ["1"]
    };

    random_seed_put_get_roundtrip_size => {
        "program t\ninteger :: n, seed(4)\ncall random_seed(size=n)\ncall random_seed(put=seed)\ncall random_seed(get=seed)\nprint *, merge(1, 0, n == 4)\nend program t\n",
        ["1"]
    };

    random_reseed_same_value_reproducible => {
        "program t\ninteger :: seed(1) = [12345]\nreal :: r1, r2\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_seed(put=seed)\ncall random_number(r2)\nprint *, merge(1, 0, r1 == r2)\nend program t\n",
        ["1"]
    };

    random_different_seeds_may_differ => {
        "program t\ninteger :: s1(1) = [1], s2(1) = [2]\nreal :: r1, r2\ncall random_seed(put=s1)\ncall random_number(r1)\ncall random_seed(put=s2)\ncall random_number(r2)\nprint *, merge(1, 0, r1 /= r2 .or. r1 == r2)\nend program t\n",
        ["1"]
    };

    random_number_after_seed_not_constant_zero => {
        "program t\ninteger :: seed(1) = [99]\nreal :: r\ncall random_seed(put=seed)\ncall random_number(r)\nprint *, merge(1, 0, r /= 0.0 .or. r == 0.0)\nend program t\n",
        ["1"]
    };

    random_array_fill_changes_sum => {
        "program t\nreal :: a(10)\ncall random_number(a)\nprint *, merge(1, 0, sum(a) >= 0.0)\nend program t\n",
        ["1"]
    };

    random_two_calls_without_reseed => {
        "program t\ninteger :: seed(1) = [7]\nreal :: r1, r2\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_number(r2)\nprint *, merge(1, 0, r1 >= 0.0 .and. r2 >= 0.0)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_1 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 1)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_2 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 2)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_3 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 3)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_4 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 4)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_5 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 5)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_6 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 6)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_7 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 7)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_8 => {
        "program t\ninteger :: sz\ncall random_seed(size=sz)\nprint *, merge(1, 0, sz >= 8)\nend program t\n",
        ["1"]
    };

    random_do_loop_ten_values_bounded => {
        "program t\nreal :: r\ninteger :: i, ok\nok = 1\ndo i = 1, 10\n  call random_number(r)\n  if (r < 0.0 .or. r >= 1.0) ok = 0\nend do\nprint *, ok\nend program t\n",
        ["1"]
    };

    random_do_loop_reseed_each_iter => {
        "program t\ninteger :: seed(1)\nreal :: r\ninteger :: i\nseed(1) = 100\ndo i = 1, 3\n  call random_seed(put=seed)\n  call random_number(r)\nend do\nprint *, merge(1, 0, r >= 0.0)\nend program t\n",
        ["1"]
    };

    random_array_section_fill => {
        "program t\nreal :: a(6)\ncall random_number(a(2:5))\nprint *, merge(1, 0, all(a(2:5) >= 0.0 .and. a(2:5) < 1.0))\nend program t\n",
        ["1"]
    };

    random_seed_default_put => {
        "program t\ncall random_seed()\nprint *, 1\nend program t\n",
        ["1"]
    };

    random_number_used_in_merge => {
        "program t\nreal :: r\ncall random_number(r)\nprint *, merge(1, 0, r < 1.0)\nend program t\n",
        ["1"]
    };

    random_number_assign_to_double => {
        "program t\ndouble precision :: r\ncall random_number(r)\nprint *, merge(1, 0, r >= 0.0d0 .and. r < 1.0d0)\nend program t\n",
        ["1"]
    };

    random_seed_multiple_values_put => {
        "program t\ninteger :: seed(4) = [1,2,3,4]\ncall random_seed(put=seed)\nprint *, 1\nend program t\n",
        ["1"]
    };

    random_reseed_restores_first_draw => {
        "program t\ninteger :: seed(1) = [555]\nreal :: r1, r2, r3\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_number(r2)\ncall random_seed(put=seed)\ncall random_number(r3)\nprint *, merge(1, 0, r3 == r1)\nend program t\n",
        ["1"]
    };

    random_compare_two_arrays_same_seed => {
        "program t\ninteger :: seed(1) = [8080]\nreal :: a(3), b(3)\ncall random_seed(put=seed)\ncall random_number(a)\ncall random_seed(put=seed)\ncall random_number(b)\nprint *, merge(1, 0, a(1) == b(1) .and. a(2) == b(2))\nend program t\n",
        ["1"]
    };

    random_sum_statistic_bounded => {
        "program t\nreal :: r(20)\ncall random_number(r)\nprint *, merge(1, 0, sum(r) < 20.0)\nend program t\n",
        ["1"]
    };

    random_max_less_than_one => {
        "program t\nreal :: r(8)\ncall random_number(r)\nprint *, merge(1, 0, maxval(r) < 1.0)\nend program t\n",
        ["1"]
    };

    random_min_non_negative => {
        "program t\nreal :: r(8)\ncall random_number(r)\nprint *, merge(1, 0, minval(r) >= 0.0)\nend program t\n",
        ["1"]
    };

    random_if_branch_in_range => {
        "program t\nreal :: r\ncall random_number(r)\nif (r >= 0.0 .and. r < 1.0) then\nprint *, 'ok'\nelse\nprint *, 'bad'\nend if\nend program t\n",
        ["ok"]
    };

    random_seed_get_after_put => {
        "program t\ninteger :: s1(2) = [11, 22], s2(2)\ncall random_seed(put=s1)\ncall random_seed(get=s2)\nprint *, merge(1, 0, s2(1) == 11 .and. s2(2) == 22)\nend program t\n",
        ["1"]
    };

    random_number_in_expression => {
        "program t\nreal :: r, x\ncall random_number(r)\nx = r * 0.0 + 0.5\nprint *, merge(1, 0, x == 0.5)\nend program t\n",
        ["1"]
    };

    random_seed_size_at_least_one => {
        "program t\ninteger :: n\ncall random_seed(size=n)\nprint *, merge(1, 0, n >= 1)\nend program t\n",
        ["1"]
    };

    random_three_reseed_identical => {
        "program t\ninteger :: seed(1) = [31415]\nreal :: r1, r2, r3\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_seed(put=seed)\ncall random_number(r2)\ncall random_seed(put=seed)\ncall random_number(r3)\nprint *, merge(1, 0, r1 == r2 .and. r2 == r3)\nend program t\n",
        ["1"]
    };

    random_array_index_single_element => {
        "program t\nreal :: a(1)\ncall random_number(a(1))\nprint *, merge(1, 0, a(1) >= 0.0 .and. a(1) < 1.0)\nend program t\n",
        ["1"]
    };

    random_large_array_hundred => {
        "program t\nreal :: r(100)\ncall random_number(r)\nprint *, merge(1, 0, count(r >= 0.0 .and. r < 1.0) == 100)\nend program t\n",
        ["1"]
    };

    random_seed_zero_allowed => {
        "program t\ninteger :: seed(1) = [0]\ncall random_seed(put=seed)\nprint *, 1\nend program t\n",
        ["1"]
    };

    random_seed_negative_values_compile => {
        "program t\ninteger :: seed(1) = [-1]\ncall random_seed(put=seed)\nprint *, 1\nend program t\n",
        ["1"]
    };

    random_number_twice_then_reseed_first => {
        "program t\ninteger :: seed(1) = [42]\nreal :: r1, r2, r3\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_number(r2)\ncall random_seed(put=seed)\ncall random_number(r3)\nprint *, merge(1, 0, r3 == r1)\nend program t\n",
        ["1"]
    };

    random_in_do_while_loop => {
        "program t\nreal :: r\ninteger :: n\nn = 0\ndo while (n < 5)\n  call random_number(r)\n  n = n + 1\nend do\nprint *, merge(1, 0, r >= 0.0)\nend program t\n",
        ["1"]
    };

    random_with_selected_real_kind => {
        "program t\ninteger, parameter :: sp = selected_real_kind(6)\nreal(sp) :: r\ncall random_number(r)\nprint *, merge(1, 0, r >= 0.0)\nend program t\n",
        ["1"]
    };

    random_seed_size_matches_put_array => {
        "program t\ninteger :: n, seed(3) = [5,6,7], got(3)\ncall random_seed(size=n)\ncall random_seed(put=seed)\ncall random_seed(get=got)\nprint *, merge(1, 0, n == 3 .and. got(2) == 6)\nend program t\n",
        ["1"]
    };

    random_number_print_bounded_flag => {
        "program t\nreal :: r\ncall random_number(r)\nprint *, merge(1, 0, r >= 0.0)\nend program t\n",
        ["1"]
    };

    random_pairwise_reseed_equality => {
        "program t\ninteger :: seed(1) = [2024]\nreal :: r1, r2\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_seed(put=seed)\ncall random_number(r2)\nprint *, merge(1, 0, abs(r1 - r2) == 0.0)\nend program t\n",
        ["1"]
    };

    random_fill_then_check_any_positive => {
        "program t\nreal :: r(50)\ncall random_number(r)\nprint *, merge(1, 0, any(r > 0.0) .or. all(r == 0.0))\nend program t\n",
        ["1"]
    };

    random_seed_large_values => {
        "program t\ninteger :: seed(1) = [999999]\ncall random_seed(put=seed)\nprint *, 1\nend program t\n",
        ["1"]
    };

    random_contained_in_select_case => {
        "program t\nreal :: r\ncall random_number(r)\nselect case (merge(1, 0, r < 1.0))\ncase (1)\nprint *, 'in'\ncase default\nprint *, 'out'\nend select\nend program t\n",
        ["in"]
    };

    random_number_vector_norm_finite => {
        "program t\nreal :: v(3)\ncall random_number(v)\nprint *, merge(1, 0, sqrt(v(1)**2 + v(2)**2 + v(3)**2) >= 0.0)\nend program t\n",
        ["1"]
    };

    random_seed_alternating_put => {
        "program t\ninteger :: s1(1) = [10], s2(1) = [20]\nreal :: r\ncall random_seed(put=s1)\ncall random_number(r)\ncall random_seed(put=s2)\ncall random_number(r)\ncall random_seed(put=s1)\ncall random_number(r)\nprint *, merge(1, 0, r >= 0.0)\nend program t\n",
        ["1"]
    };

    random_array_assign_from_previous => {
        "program t\nreal :: a(2), b(2)\ncall random_number(a)\nb = a\nprint *, merge(1, 0, b(1) == a(1) .and. b(2) == a(2))\nend program t\n",
        ["1"]
    };

    random_number_floor_is_zero_or_more => {
        "program t\nreal :: r\ncall random_number(r)\nprint *, merge(1, 0, int(r) >= 0)\nend program t\n",
        ["1"]
    };

    random_count_below_half => {
        "program t\nreal :: r(30)\ninteger :: c\ncall random_number(r)\nc = count(r < 0.5)\nprint *, merge(1, 0, c >= 0 .and. c <= 30)\nend program t\n",
        ["1"]
    };

    random_seed_identity_after_get_put => {
        "program t\ninteger :: a(2), b(2)\ncall random_seed(get=a)\ncall random_seed(put=a)\ncall random_seed(get=b)\nprint *, merge(1, 0, a(1) == b(1) .and. a(2) == b(2))\nend program t\n",
        ["1"]
    };

    random_reseed_after_many_draws => {
        "program t\ninteger :: seed(1) = [77]\nreal :: r, first\ninteger :: i\ncall random_seed(put=seed)\ncall random_number(first)\ndo i = 1, 20\n  call random_number(r)\nend do\ncall random_seed(put=seed)\ncall random_number(r)\nprint *, merge(1, 0, r == first)\nend program t\n",
        ["1"]
    };

    random_boolean_from_compare => {
        "program t\nreal :: r\ncall random_number(r)\nprint *, r < 1.0\nend program t\n",
        ["true"]
    };

    random_seed_get_without_prior_put => {
        "program t\ninteger :: seed(8)\ncall random_seed(get=seed)\nprint *, merge(1, 0, size(seed) >= 1)\nend program t\n",
        ["1"]
    };

    random_draw_advances_without_reseed => {
        "program t\ninteger :: seed(1) = [13]\nreal :: r1, r2\ncall random_seed(put=seed)\ncall random_number(r1)\ncall random_number(r2)\nprint *, merge(1, 0, r1 >= 0.0 .and. r2 >= 0.0 .and. r1 /= r2 .or. r1 == r2)\nend program t\n",
        ["1"]
    };

    random_number_in_real_function => {
        "program t\nprint *, merge(1, 0, draw() >= 0.0 .and. draw() < 1.0)\ncontains\nfunction draw() result(r)\nreal :: r\ncall random_number(r)\nend function draw\nend program t\n",
        ["1"]
    };

    random_seed_put_then_number_then_get => {
        "program t\ninteger :: s(1) = [888], g(1)\nreal :: r\ncall random_seed(put=s)\ncall random_number(r)\ncall random_seed(get=g)\nprint *, merge(1, 0, g(1) == 888 .and. r >= 0.0)\nend program t\n",
        ["1"]
    };

}

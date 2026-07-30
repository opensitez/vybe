//! Extended DO CONCURRENT: index triples, stride/mask, LOCAL/SHARED locality,
//! reduction-style fill-then-sum, and named concurrent loops.
//! Distinct from `test_do_loop_extended.rs` and `test_fortran2008.rs` (basic concurrent).

fortran_cases! {
    // ── Simple index ranges ──────────────────────────────────────────

    do_concurrent_fill_squares => {
        "program t\ninteger :: a(8)\ndo concurrent (i = 1:8)\na(i) = i * i\nend do\nprint *, a(5)\nend program t\n",
        ["25"]
    };

    do_concurrent_fill_linear => {
        "program t\ninteger :: a(6)\ndo concurrent (i = 1:6)\na(i) = i * 10\nend do\nprint *, a(4)\nend program t\n",
        ["40"]
    };

    do_concurrent_reverse_index_order => {
        "program t\ninteger :: a(5)\ndo concurrent (i = 1:5)\na(i) = 6 - i\nend do\nprint *, a(1)\nprint *, a(5)\nend program t\n",
        ["5", "1"]
    };

    do_concurrent_single_element => {
        "program t\ninteger :: a(1)\ndo concurrent (i = 1:1)\na(i) = 99\nend do\nprint *, a(1)\nend program t\n",
        ["99"]
    };

    do_concurrent_identity_map => {
        "program t\ninteger :: a(7)\ndo concurrent (i = 1:7)\na(i) = i\nend do\nprint *, a(7)\nend program t\n",
        ["7"]
    };

    // ── Stride variants ────────────────────────────────────────────────

    do_concurrent_stride_two => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10:2)\na(i) = i\nend do\nprint *, a(7)\nend program t\n",
        ["7"]
    };

    do_concurrent_stride_three => {
        "program t\ninteger :: a(12)\na = 0\ndo concurrent (i = 1:12:3)\na(i) = i\nend do\nprint *, a(10)\nend program t\n",
        ["10"]
    };

    do_concurrent_stride_skip_odds => {
        "program t\ninteger :: a(8)\na = 0\ndo concurrent (i = 2:8:2)\na(i) = i\nend do\nprint *, a(4)\nend program t\n",
        ["4"]
    };

    do_concurrent_stride_unfilled_slot => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10:3)\na(i) = 1\nend do\nprint *, a(2)\nend program t\n",
        ["0"]
    };

    do_concurrent_stride_sum_filled => {
        "program t\ninteger :: a(10), s, k\na = 0\ndo concurrent (i = 1:10:2)\na(i) = i\nend do\ns = 0\ndo k = 1, 10\ns = s + a(k)\nend do\nprint *, s\nend program t\n",
        ["25"]
    };

    // ── Two-dimensional indices ────────────────────────────────────────

    do_concurrent_2d_row_major => {
        "program t\ninteger :: m(3,4)\nm = 0\ndo concurrent (i = 1:3, j = 1:4)\nm(i,j) = i * 10 + j\nend do\nprint *, m(2,3)\nend program t\n",
        ["23"]
    };

    do_concurrent_2d_diagonal => {
        "program t\ninteger :: m(4,4)\nm = 0\ndo concurrent (i = 1:4, j = 1:4)\nif (i == j) m(i,j) = i\nend do\nprint *, m(3,3)\nend program t\n",
        ["3"]
    };

    do_concurrent_2d_off_diagonal_sum => {
        "program t\ninteger :: m(3,3)\nm = 0\ndo concurrent (i = 1:3, j = 1:3)\nif (i /= j) m(i,j) = 1\nend do\nprint *, sum(m)\nend program t\n",
        ["6"]
    };

    do_concurrent_2d_product_table => {
        "program t\ninteger :: m(5,5)\ndo concurrent (i = 1:5, j = 1:5)\nm(i,j) = i * j\nend do\nprint *, m(3,4)\nend program t\n",
        ["12"]
    };

    do_concurrent_2d_corner_elements => {
        "program t\ninteger :: m(2,2)\ndo concurrent (i = 1:2, j = 1:2)\nm(i,j) = i + j\nend do\nprint *, m(1,1)\nprint *, m(2,2)\nend program t\n",
        ["2", "4"]
    };

    // ── Three-dimensional indices ──────────────────────────────────────

    do_concurrent_3d_fill => {
        "program t\ninteger :: a(2,2,2)\ndo concurrent (i = 1:2, j = 1:2, k = 1:2)\na(i,j,k) = i + j + k\nend do\nprint *, a(1,2,1)\nend program t\n",
        ["4"]
    };

    do_concurrent_3d_count_cells => {
        "program t\ninteger :: a(3,3,3), c, ii, jj, kk\na = 1\nc = 0\ndo ii = 1, 3\ndo jj = 1, 3\ndo kk = 1, 3\nc = c + a(ii,jj,kk)\nend do\nend do\nend do\nprint *, c\nend program t\n",
        ["27"]
    };

    do_concurrent_3d_layer_slice => {
        "program t\ninteger :: a(2,2,3)\na = 0\ndo concurrent (i = 1:2, j = 1:2, k = 1:3)\nif (k == 2) a(i,j,k) = i + j\nend do\nprint *, a(1,1,2)\nend program t\n",
        ["2"]
    };

    // ── Mask conditions ────────────────────────────────────────────────

    do_concurrent_mask_even_only => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10, mod(i,2) == 0)\na(i) = i\nend do\nprint *, a(6)\nend program t\n",
        ["6"]
    };

    do_concurrent_mask_odd_only => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10, mod(i,2) /= 0)\na(i) = i * 2\nend do\nprint *, a(5)\nend program t\n",
        ["10"]
    };

    do_concurrent_mask_greater_than_five => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10, i > 5)\na(i) = 1\nend do\nprint *, sum(a)\nend program t\n",
        ["5"]
    };

    do_concurrent_mask_less_equal_three => {
        "program t\ninteger :: a(10)\na = 0\ndo concurrent (i = 1:10, i <= 3)\na(i) = i\nend do\nprint *, a(4)\nend program t\n",
        ["0"]
    };

    do_concurrent_mask_2d_upper_triangle => {
        "program t\ninteger :: m(4,4)\nm = 0\ndo concurrent (i = 1:4, j = 1:4, j >= i)\nm(i,j) = 1\nend do\nprint *, sum(m)\nend program t\n",
        ["10"]
    };

    do_concurrent_mask_prime_slots => {
        "program t\ninteger :: a(12)\na = 0\ndo concurrent (i = 2:12, mod(i,2)/=0 .or. i==2)\na(i) = 1\nend do\nprint *, a(3)\nend program t\n",
        ["1"]
    };

    do_concurrent_mask_mod_three_zero => {
        "program t\ninteger :: a(15)\na = 0\ndo concurrent (i = 1:15, mod(i,3) == 0)\na(i) = i\nend do\nprint *, a(9)\nend program t\n",
        ["9"]
    };

    // ── LOCAL locality clause ──────────────────────────────────────────

    do_concurrent_local_temp_doubles => {
        "program t\ninteger :: src(5), dst(5)\nsrc = [1, 2, 3, 4, 5]\ndo concurrent (i = 1:5) local(tmp)\ninteger :: tmp\ntmp = src(i) * 2\ndst(i) = tmp\nend do\nprint *, dst(3)\nend program t\n",
        ["6"]
    };

    do_concurrent_local_accumulate_square => {
        "program t\ninteger :: a(6)\ndo concurrent (i = 1:6) local(sq)\ninteger :: sq\nsq = i * i\na(i) = sq\nend do\nprint *, a(4)\nend program t\n",
        ["16"]
    };

    do_concurrent_local_real_temp => {
        "program t\nreal :: a(4), b(4)\nb = [1.0, 2.0, 3.0, 4.0]\ndo concurrent (i = 1:4) local(t)\nreal :: t\nt = b(i) * 2.5\na(i) = t\nend do\nprint *, int(a(2))\nend program t\n",
        ["5"]
    };

    do_concurrent_local_with_mask => {
        "program t\ninteger :: a(8)\na = 0\ndo concurrent (i = 1:8, mod(i,2)==0) local(v)\ninteger :: v\nv = i * 3\na(i) = v\nend do\nprint *, a(4)\nend program t\n",
        ["12"]
    };

    do_concurrent_local_2d => {
        "program t\ninteger :: m(2,2)\ndo concurrent (i = 1:2, j = 1:2) local(prod)\ninteger :: prod\nprod = i * j\nm(i,j) = prod\nend do\nprint *, m(2,2)\nend program t\n",
        ["4"]
    };

    // ── SHARED locality clause ─────────────────────────────────────────

    do_concurrent_shared_scale_factor => {
        "program t\ninteger :: a(5), factor\nfactor = 4\ndo concurrent (i = 1:5) shared(factor)\na(i) = i * factor\nend do\nprint *, a(3)\nend program t\n",
        ["12"]
    };

    do_concurrent_shared_offset => {
        "program t\ninteger :: a(6), base\nbase = 10\ndo concurrent (i = 1:6) shared(base)\na(i) = base + i\nend do\nprint *, a(2)\nend program t\n",
        ["12"]
    };

    do_concurrent_shared_multiplier_real => {
        "program t\nreal :: a(4), mult\nmult = 2.0\ndo concurrent (i = 1:4) shared(mult)\na(i) = real(i) * mult\nend do\nprint *, int(a(3))\nend program t\n",
        ["6"]
    };

    do_concurrent_shared_with_mask => {
        "program t\ninteger :: a(10), step\nstep = 5\na = 0\ndo concurrent (i = 1:10, mod(i,step)==0) shared(step)\na(i) = i\nend do\nprint *, a(5)\nend program t\n",
        ["5"]
    };

    // ── Reduction-style fill then sequential sum ───────────────────────

    do_concurrent_fill_then_sum_all => {
        "program t\ninteger :: a(10), s, k\ndo concurrent (i = 1:10)\na(i) = i\nend do\ns = 0\ndo k = 1, 10\ns = s + a(k)\nend do\nprint *, s\nend program t\n",
        ["55"]
    };

    do_concurrent_fill_squares_then_sum => {
        "program t\ninteger :: a(5), s, k\ndo concurrent (i = 1:5)\na(i) = i * i\nend do\ns = 0\ndo k = 1, 5\ns = s + a(k)\nend do\nprint *, s\nend program t\n",
        ["55"]
    };

    do_concurrent_count_nonzero_mask => {
        "program t\ninteger :: a(10), c, k\na = 0\ndo concurrent (i = 1:10, mod(i,2)==0)\na(i) = 1\nend do\nc = 0\ndo k = 1, 10\nif (a(k) /= 0) c = c + 1\nend do\nprint *, c\nend program t\n",
        ["5"]
    };

    do_concurrent_max_via_sequential_scan => {
        "program t\ninteger :: a(8), mx, k\ndo concurrent (i = 1:8)\na(i) = i * 3\nend do\nmx = a(1)\ndo k = 2, 8\nif (a(k) > mx) mx = a(k)\nend do\nprint *, mx\nend program t\n",
        ["24"]
    };

    do_concurrent_product_via_sequential => {
        "program t\ninteger :: a(4), p, k\ndo concurrent (i = 1:4)\na(i) = i + 1\nend do\np = 1\ndo k = 1, 4\np = p * a(k)\nend do\nprint *, p\nend program t\n",
        ["120"]
    };

    // ── Real and logical arrays ────────────────────────────────────────

    do_concurrent_real_array => {
        "program t\nreal :: r(5)\ndo concurrent (i = 1:5)\nr(i) = real(i) * 1.5\nend do\nprint *, int(r(4))\nend program t\n",
        ["6"]
    };

    do_concurrent_logical_mask_array => {
        "program t\nlogical :: flags(6)\ndo concurrent (i = 1:6)\nflags(i) = mod(i,2) == 0\nend do\nprint *, flags(4)\nend program t\n",
        ["true"]
    };

    do_concurrent_logical_all_false => {
        "program t\nlogical :: flags(4)\ndo concurrent (i = 1:4)\nflags(i) = .false.\nend do\nprint *, flags(1)\nend program t\n",
        ["false"]
    };

    // ── Named DO CONCURRENT ────────────────────────────────────────────

    do_concurrent_named_fill => {
        "program t\ninteger :: a(5)\nfill: do concurrent (i = 1:5)\na(i) = i * 10\nend do fill\nprint *, a(3)\nend program t\n",
        ["30"]
    };

    do_concurrent_named_with_mask => {
        "program t\ninteger :: a(8)\na = 0\neven: do concurrent (i = 1:8, mod(i,2)==0)\na(i) = i\nend do even\nprint *, a(6)\nend program t\n",
        ["6"]
    };

    // ── Inside block and sequential concurrent loops ───────────────────

    do_concurrent_inside_block => {
        "program t\nblock\ninteger :: a(4)\ndo concurrent (i = 1:4)\na(i) = i + 1\nend do\nprint *, a(4)\nend block\nend program t\n",
        ["5"]
    };

    do_concurrent_two_sequential_loops => {
        "program t\ninteger :: a(5), b(5)\ndo concurrent (i = 1:5)\na(i) = i\nend do\ndo concurrent (i = 1:5)\nb(i) = a(i) * 2\nend do\nprint *, b(3)\nend program t\n",
        ["6"]
    };

    do_concurrent_overwrite_prior => {
        "program t\ninteger :: a(4)\ndo concurrent (i = 1:4)\na(i) = i\nend do\ndo concurrent (i = 1:4)\na(i) = a(i) + 10\nend do\nprint *, a(2)\nend program t\n",
        ["12"]
    };

    // ── Edge cases and combined forms ──────────────────────────────────

    do_concurrent_zero_length_range => {
        "program t\ninteger :: a(5)\na = 1\ndo concurrent (i = 6:10)\na(1) = 99\nend do\nprint *, a(1)\nend program t\n",
        ["1"]
    };

    do_concurrent_large_index => {
        "program t\ninteger, allocatable :: a(:)\nallocate(a(100))\na = 0\ndo concurrent (i = 1:100)\na(i) = i\nend do\nprint *, a(100)\nend program t\n",
        ["100"]
    };

    do_concurrent_negative_values_via_expr => {
        "program t\ninteger :: a(5)\ndo concurrent (i = 1:5)\na(i) = i - 3\nend do\nprint *, a(1)\nprint *, a(5)\nend program t\n",
        ["-2", "2"]
    };

    do_concurrent_2d_stride_mask => {
        "program t\ninteger :: m(4,4)\nm = 0\ndo concurrent (i = 1:4:2, j = 1:4:2)\nm(i,j) = i + j\nend do\nprint *, m(1,1)\nprint *, m(3,3)\nend program t\n",
        ["2", "6"]
    };

    do_concurrent_local_shared_combined => {
        "program t\ninteger :: a(4), scale\nscale = 3\ndo concurrent (i = 1:4) shared(scale) local(t)\ninteger :: t\nt = i * scale\na(i) = t\nend do\nprint *, a(4)\nend program t\n",
        ["12"]
    };

    do_concurrent_character_array => {
        "program t\ncharacter(len=1) :: chars(3)\ndo concurrent (i = 1:3)\nchars(i) = achar(96 + i)\nend do\nprint *, chars(2)\nend program t\n",
        ["b"]
    };

    do_concurrent_copy_from_source => {
        "program t\ninteger :: src(6), dst(6)\nsrc = [10, 20, 30, 40, 50, 60]\ndo concurrent (i = 1:6)\ndst(i) = src(i)\nend do\nprint *, dst(5)\nend program t\n",
        ["50"]
    };

    do_concurrent_reverse_values => {
        "program t\ninteger :: a(5), b(5)\ndo concurrent (i = 1:5)\na(i) = i\nend do\ndo concurrent (i = 1:5)\nb(i) = a(6 - i)\nend do\nprint *, b(1)\nend program t\n",
        ["5"]
    };

    do_concurrent_sum_diagonal_2d => {
        "program t\ninteger :: m(5,5), s, k\nm = 0\ndo concurrent (i = 1:5, j = 1:5)\nif (i == j) m(i,j) = i\nend do\ns = 0\ndo k = 1, 5\ns = s + m(k,k)\nend do\nprint *, s\nend program t\n",
        ["15"]
    };

    do_concurrent_checkerboard => {
        "program t\ninteger :: m(4,4)\nm = 0\ndo concurrent (i = 1:4, j = 1:4, mod(i+j,2)==0)\nm(i,j) = 1\nend do\nprint *, sum(m)\nend program t\n",
        ["8"]
    };

    do_concurrent_triple_index_product => {
        "program t\ninteger :: v\nv = 0\ndo concurrent (i = 1:2, j = 1:2, k = 1:2)\nv = v + 1\nend do\nprint *, v\nend program t\n",
        ["8"]
    };

    do_concurrent_mask_combined_and => {
        "program t\ninteger :: a(12)\na = 0\ndo concurrent (i = 1:12, i > 2 .and. i < 10)\na(i) = 1\nend do\nprint *, sum(a)\nend program t\n",
        ["7"]
    };

    do_concurrent_fill_descending_index => {
        "program t\ninteger :: a(5)\ndo concurrent (i = 1:5)\na(i) = 6 - i\nend do\nprint *, a(3)\nend program t\n",
        ["3"]
    };

    do_concurrent_variable_bounds => {
        "program t\ninteger :: a(9), n\nn = 9\ndo concurrent (i = 1:n)\na(i) = i + 1\nend do\nprint *, a(5)\nend program t\n",
        ["6"]
    };

    do_concurrent_variable_stride => {
        "program t\ninteger :: a(10), stride\na = 0\nstride = 3\ndo concurrent (i = 1:10:stride)\na(i) = i\nend do\nprint *, a(7)\nend program t\n",
        ["7"]
    };

    do_concurrent_variable_mask_limit => {
        "program t\ninteger :: a(10), limit\na = 0\nlimit = 6\ndo concurrent (i = 1:10, i <= limit)\na(i) = 1\nend do\nprint *, sum(a)\nend program t\n",
        ["6"]
    };

    do_concurrent_bounds_mutation_is_ignored => {
        "program t\ninteger :: a(12), n\nn = 12\na = 0\ndo concurrent (i = 1:n)\na(i) = 1\nif (i == 4) n = 3\nend do\nprint *, sum(a)\nend program t\n",
        ["12"]
    };

    do_concurrent_named_with_variable_bounds => {
        "program t\ninteger :: a(7), n\nn = 7\nfill: do concurrent (i = 1:n)\na(i) = 2 * i\nend do fill\nprint *, a(4)\nend program t\n",
        ["8"]
    };
}

//! Extended SELECT TYPE and SELECT RANK: intrinsic type guards (integer, real,
//! character, logical), CLASS IS / CLASS DEFAULT, and rank(0..3) branches.
//! Distinct from `test_derived_type_oop_extended.rs` (OOP guards compile-only),
//! `test_fortran2003.rs`, and `test_fortran2018.rs` (basic rank forms).

fortran_cases! {
    // ── SELECT TYPE: TYPE IS (integer) ───────────────────────────────

    select_type_integer_guard_matches => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 42\nselect type(val)\ntype is (integer)\nprint *, val\ntype is (real)\nprint *, 0\nclass default\nprint *, -1\nend select\nend program t\n",
        ["42"]
    };

    select_type_integer_not_real => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 7\nselect type(val)\ntype is (real)\nprint *, 0\ntype is (integer)\nprint *, val\nend select\nend program t\n",
        ["7"]
    };

    select_type_integer_negative_value => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = -15\nselect type(val)\ntype is (integer)\nprint *, val\nclass default\nprint *, 0\nend select\nend program t\n",
        ["-15"]
    };

    select_type_integer_zero => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 0\nselect type(val)\ntype is (integer)\nprint *, val\nend select\nend program t\n",
        ["0"]
    };

    select_type_integer_large => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 1000000\nselect type(val)\ntype is (integer)\nprint *, val / 1000000\nend select\nend program t\n",
        ["1"]
    };

    // ── SELECT TYPE: TYPE IS (real) ────────────────────────────────────

    select_type_real_guard_matches => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 3.5\nselect type(val)\ntype is (real)\nprint *, int(val)\ntype is (integer)\nprint *, 0\nend select\nend program t\n",
        ["3"]
    };

    select_type_real_not_integer => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 2.7\nselect type(val)\ntype is (integer)\nprint *, 0\ntype is (real)\nprint *, int(val * 10.0)\nend select\nend program t\n",
        ["27"]
    };

    select_type_real_negative => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = -4.0\nselect type(val)\ntype is (real)\nprint *, int(abs(val))\nend select\nend program t\n",
        ["4"]
    };

    select_type_real_fractional => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 0.25\nselect type(val)\ntype is (real)\nprint *, int(val * 100.0)\nend select\nend program t\n",
        ["25"]
    };

    // ── SELECT TYPE: TYPE IS (character) ───────────────────────────────

    select_type_character_guard_matches => {
        "program t\nclass(*), allocatable :: val\nallocate(character(len=5) :: val)\nval = 'hello'\nselect type(val)\ntype is (character(len=*))\nprint *, trim(val)\ntype is (integer)\nprint *, 'no'\nend select\nend program t\n",
        ["hello"]
    };

    select_type_character_short => {
        "program t\nclass(*), allocatable :: val\nallocate(character(len=3) :: val)\nval = 'abc'\nselect type(val)\ntype is (character(len=*))\nprint *, len_trim(val)\nend select\nend program t\n",
        ["3"]
    };

    select_type_character_empty => {
        "program t\nclass(*), allocatable :: val\nallocate(character(len=1) :: val)\nval = 'x'\nselect type(val)\ntype is (character(len=*))\nprint *, val\nend select\nend program t\n",
        ["x"]
    };

    select_type_character_not_integer => {
        "program t\nclass(*), allocatable :: val\nallocate(character(len=4) :: val)\nval = 'test'\nselect type(val)\ntype is (integer)\nprint *, 0\ntype is (character(len=*))\nprint *, trim(val)\nend select\nend program t\n",
        ["test"]
    };

    // ── SELECT TYPE: TYPE IS (logical) ─────────────────────────────────

    select_type_logical_true => {
        "program t\nclass(*), allocatable :: val\nallocate(logical :: val)\nval = .true.\nselect type(val)\ntype is (logical)\nprint *, val\nend select\nend program t\n",
        ["true"]
    };

    select_type_logical_false => {
        "program t\nclass(*), allocatable :: val\nallocate(logical :: val)\nval = .false.\nselect type(val)\ntype is (logical)\nprint *, val\nend select\nend program t\n",
        ["false"]
    };

    select_type_logical_not_integer => {
        "program t\nclass(*), allocatable :: val\nallocate(logical :: val)\nval = .true.\nselect type(val)\ntype is (integer)\nprint *, 0\ntype is (logical)\nprint *, val\nend select\nend program t\n",
        ["true"]
    };

    // ── SELECT TYPE: CLASS DEFAULT ─────────────────────────────────────

    select_type_class_default_integer => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 9\nselect type(val)\ntype is (real)\nprint *, 0\nclass default\nprint *, val\nend select\nend program t\n",
        ["9"]
    };

    select_type_class_default_real => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 6.0\nselect type(val)\ntype is (integer)\nprint *, 0\nclass default\nprint *, int(val)\nend select\nend program t\n",
        ["6"]
    };

    select_type_class_default_only => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 55\nselect type(val)\nclass default\nprint *, val\nend select\nend program t\n",
        ["55"]
    };

    select_type_no_match_hits_default => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 1.0\nselect type(val)\ntype is (integer)\nprint *, 1\ntype is (character(len=*))\nprint *, 2\nclass default\nprint *, 3\nend select\nend program t\n",
        ["3"]
    };

    // ── SELECT TYPE: derived TYPE IS / CLASS IS ────────────────────────

    select_type_dtype_base_type_is => {
        "program t\ntype :: Base\ninteger :: id = 5\nend type Base\nclass(Base), allocatable :: obj\nallocate(Base :: obj)\nselect type(obj)\ntype is (Base)\nprint *, obj%id\nclass default\nprint *, 0\nend select\nend program t\n",
        ["5"]
    };

    select_type_dtype_child_class_is => {
        "program t\ntype :: Base\ninteger :: id = 1\nend type Base\ntype, extends(Base) :: Child\ninteger :: extra = 9\nend type Child\nclass(Base), allocatable :: obj\nallocate(Child :: obj)\nselect type(obj)\nclass is (Child)\nprint *, obj%extra\ntype is (Base)\nprint *, obj%id\nend select\nend program t\n",
        ["9"]
    };

    select_type_dtype_child_type_is_base => {
        "program t\ntype :: Base\ninteger :: id = 3\nend type Base\ntype, extends(Base) :: Child\ninteger :: extra = 7\nend type Child\nclass(Base), allocatable :: obj\nallocate(Base :: obj)\nselect type(obj)\ntype is (Base)\nprint *, obj%id\nclass is (Child)\nprint *, obj%extra\nend select\nend program t\n",
        ["3"]
    };

    select_type_dtype_class_default_base => {
        "program t\ntype :: Root\ninteger :: tag = 2\nend type Root\nclass(Root), allocatable :: node\nallocate(Root :: node)\nselect type(node)\nclass is (Root)\nprint *, node%tag\nclass default\nprint *, 0\nend select\nend program t\n",
        ["2"]
    };

    select_type_dtype_two_level_hierarchy => {
        "program t\ntype :: A\ninteger :: x = 1\nend type A\ntype, extends(A) :: B\ninteger :: y = 2\nend type B\ntype, extends(B) :: C\ninteger :: z = 3\nend type C\nclass(A), allocatable :: obj\nallocate(C :: obj)\nselect type(obj)\nclass is (C)\nprint *, obj%z\nclass is (B)\nprint *, obj%y\ntype is (A)\nprint *, obj%x\nend select\nend program t\n",
        ["3"]
    };

    select_type_dtype_class_is_before_type_is => {
        "program t\ntype :: P\ninteger :: v = 10\nend type P\ntype, extends(P) :: Q\ninteger :: w = 20\nend type Q\nclass(P), allocatable :: obj\nallocate(Q :: obj)\nselect type(obj)\nclass is (Q)\nprint *, obj%w\ntype is (P)\nprint *, obj%v\nend select\nend program t\n",
        ["20"]
    };

    // ── SELECT TYPE: unlimited polymorphic array element ───────────────

    select_type_integer_array_elem => {
        "program t\nclass(*), allocatable :: arr(:)\nallocate(integer :: arr(3))\narr = [4, 5, 6]\nselect type(arr(2))\ntype is (integer)\nprint *, arr(2)\nend select\nend program t\n",
        ["5"]
    };

    select_type_real_in_mixed_check => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 8.0\nselect type(val)\ntype is (integer)\nprint *, 1\ntype is (real)\nprint *, int(val)\ntype is (logical)\nprint *, 2\nclass default\nprint *, 3\nend select\nend program t\n",
        ["8"]
    };

    // ── SELECT RANK: rank(0) scalar ───────────────────────────────────

    select_rank_scalar_rank0 => {
        "program t\ncall tag(99)\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(0)\nprint *, x\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["99"]
    };

    select_rank_scalar_not_rank1 => {
        "program t\ncall tag(7)\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, 0\nrank(0)\nprint *, x\nend select\nend subroutine tag\nend program t\n",
        ["7"]
    };

    select_rank_real_scalar => {
        "program t\ncall tag(3.5)\ncontains\nsubroutine tag(x)\nreal, intent(in) :: x(..)\nselect rank(x)\nrank(0)\nprint *, int(x)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["3"]
    };

    // ── SELECT RANK: rank(1) vector ────────────────────────────────────

    select_rank_vector_size => {
        "program t\ncall tag([10, 20, 30, 40])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, size(x)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["4"]
    };

    select_rank_vector_first_elem => {
        "program t\ncall tag([5, 6, 7])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, x(1)\nrank(0)\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["5"]
    };

    select_rank_vector_sum => {
        "program t\ncall tag([1, 2, 3, 4])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, sum(x)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["10"]
    };

    select_rank_vector_last_elem => {
        "program t\ncall tag([2, 4, 6, 8, 10])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, x(size(x))\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["10"]
    };

    select_rank_real_vector => {
        "program t\ncall tag([1.0, 2.0, 3.0])\ncontains\nsubroutine tag(x)\nreal, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, int(sum(x))\nrank(0)\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["6"]
    };

    // ── SELECT RANK: rank(2) matrix ────────────────────────────────────

    select_rank_matrix_dims => {
        "program t\ncall tag(reshape([1,2,3,4,5,6], [2,3]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, size(x,1), size(x,2)\nrank default\nprint *, 0, 0\nend select\nend subroutine tag\nend program t\n",
        ["2", "3"]
    };

    select_rank_matrix_element => {
        "program t\ncall tag(reshape([1,2,3,4], [2,2]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, x(2,1)\nrank(1)\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["3"]
    };

    select_rank_matrix_trace => {
        "program t\ncall tag(reshape([1,0,0,2,0,0,3], [3,3]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, x(1,1) + x(2,2) + x(3,3)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["6"]
    };

    select_rank_real_matrix => {
        "program t\ncall tag(reshape([1.0, 2.0, 3.0, 4.0], [2,2]))\ncontains\nsubroutine tag(x)\nreal, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, int(x(1,1) + x(2,2))\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["5"]
    };

    // ── SELECT RANK: rank(3) and rank default ──────────────────────────

    select_rank_tensor_rank3 => {
        "program t\ncall tag(reshape([(i, i=1,8)], [2,2,2]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(3)\nprint *, size(x,1), size(x,2), size(x,3)\nrank default\nprint *, 0, 0, 0\nend select\nend subroutine tag\nend program t\n",
        ["2", "2", "2"]
    };

    select_rank_tensor_elem => {
        "program t\ncall tag(reshape([(i, i=1,8)], [2,2,2]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(3)\nprint *, x(1,1,1)\nrank(2)\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["1"]
    };

    select_rank_default_for_scalar => {
        "program t\ncall tag(42)\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, 0\nrank(2)\nprint *, 0\nrank default\nprint *, x\nend select\nend subroutine tag\nend program t\n",
        ["42"]
    };

    select_rank_default_reports_rank_fn => {
        "program t\ncall tag([1, 2])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(3)\nprint *, 0\nrank default\nprint *, rank(x)\nend select\nend subroutine tag\nend program t\n",
        ["1"]
    };

    select_rank_vector_hits_rank_default_when_no_rank1 => {
        "program t\ncall tag(5)\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, size(x)\nrank default\nprint *, rank(x)\nend select\nend subroutine tag\nend program t\n",
        ["0"]
    };

    // ── SELECT RANK: multi-branch dispatch ─────────────────────────────

    select_rank_three_way_scalar_vector_matrix => {
        "program t\ncall tag(7)\ncall tag([1,2])\ncall tag(reshape([1,2,3,4],[2,2]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(0)\nprint *, x + 100\nrank(1)\nprint *, size(x) + 200\nrank(2)\nprint *, size(x,1) * size(x,2) + 300\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["107", "202", "304"]
    };

    select_rank_module_procedure => {
        "module rankmod\ncontains\nsubroutine rows(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, size(x,1)\nrank(1)\nprint *, size(x)\nrank default\nprint *, 0\nend select\nend subroutine rows\nend module rankmod\nprogram t\nuse rankmod\ncall rows([1,2,3])\ncall rows(reshape([1,2,3,4,5,6],[2,3]))\nend program t\n",
        ["3", "2"]
    };

    // ── Combined SELECT TYPE inside rank branch ───────────────────────

    select_rank1_then_select_type_integer => {
        "program t\ncall inspect([10, 20])\ncontains\nsubroutine inspect(x)\nclass(*), intent(in) :: x(..)\nselect rank(x)\nrank(1)\nselect type(x)\ntype is (integer)\nprint *, sum(x)\nclass default\nprint *, 0\nend select\nrank default\nprint *, -1\nend select\nend subroutine inspect\nend program t\n",
        ["30"]
    };

    select_rank0_then_select_type_real => {
        "program t\ncall inspect(4.5)\ncontains\nsubroutine inspect(x)\nclass(*), intent(in) :: x(..)\nselect rank(x)\nrank(0)\nselect type(x)\ntype is (real)\nprint *, int(x)\nclass default\nprint *, 0\nend select\nrank default\nprint *, -1\nend select\nend subroutine inspect\nend program t\n",
        ["4"]
    };

    select_type_integer_vs_real_dispatch => {
        "program t\ncall show(3)\ncall show(3.0)\ncontains\nsubroutine show(val)\nclass(*), intent(in) :: val\nselect type(val)\ntype is (integer)\nprint *, val * 2\ntype is (real)\nprint *, int(val * 2.0)\nclass default\nprint *, 0\nend select\nend subroutine show\nend program t\n",
        ["6", "6"]
    };

    select_type_character_vs_integer_dispatch => {
        "program t\ncharacter(len=3) :: s = 'abc'\ncall show(s)\ncall show(5)\ncontains\nsubroutine show(val)\nclass(*), intent(in) :: val\nselect type(val)\ntype is (character(len=*))\nprint *, len_trim(val)\ntype is (integer)\nprint *, val\nclass default\nprint *, 0\nend select\nend subroutine show\nend program t\n",
        ["3", "5"]
    };

    select_rank2_integer_sum_all => {
        "program t\ncall total(reshape([1,2,3,4,5,6],[2,3]))\ncontains\nsubroutine total(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, sum(x)\nrank(1)\nprint *, sum(x)\nrank default\nprint *, x\nend select\nend subroutine total\nend program t\n",
        ["21"]
    };

    select_type_logical_in_branch => {
        "program t\nclass(*), allocatable :: val\nallocate(logical :: val)\nval = .false.\nselect type(val)\ntype is (logical)\nif (val) then\nprint *, 1\nelse\nprint *, 0\nend if\nend select\nend program t\n",
        ["0"]
    };

    select_type_reallocate_change_type => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 8\nselect type(val)\ntype is (integer)\nprint *, val\nend select\ndeallocate(val)\nallocate(real :: val)\nval = 2.5\nselect type(val)\ntype is (real)\nprint *, int(val)\nend select\nend program t\n",
        ["8", "2"]
    };

    select_rank_character_vector_len => {
        "program t\ncall tag(['a', 'bb', 'ccc'])\ncontains\nsubroutine tag(x)\ncharacter(len=*), intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, size(x)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["3"]
    };

    select_rank_logical_vector_any => {
        "program t\ncall tag([.true., .false., .true.])\ncontains\nsubroutine tag(x)\nlogical, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nif (any(x)) print *, 1\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["1"]
    };

    select_type_class_default_derived => {
        "program t\ntype :: Box\ninteger :: w = 4\nend type Box\nclass(Box), allocatable :: b\nallocate(Box :: b)\nselect type(b)\ntype is (integer)\nprint *, 0\nclass default\nprint *, b%w\nend select\nend program t\n",
        ["4"]
    };

    select_rank_assumed_rank_function_result => {
        "program t\nprint *, pick(reshape([1,2,3,4],[2,2]))\ncontains\ninteger function pick(m)\ninteger, intent(in) :: m(..)\nselect rank(m)\nrank(2)\npick = m(1,1) + m(2,2)\nrank default\npick = 0\nend select\nend function pick\nend program t\n",
        ["5"]
    };

    select_type_complex_not_matched_default => {
        "program t\nclass(*), allocatable :: val\nallocate(complex :: val)\nval = (1.0, 2.0)\nselect type(val)\ntype is (integer)\nprint *, 0\ntype is (real)\nprint *, 0\nclass default\nprint *, int(real(val) + aimag(val))\nend select\nend program t\n",
        ["3"]
    };

    select_rank_scalar_character => {
        "program t\ncall tag('z')\ncontains\nsubroutine tag(x)\ncharacter(len=*), intent(in) :: x(..)\nselect rank(x)\nrank(0)\nprint *, x\nrank default\nprint *, '?'\nend select\nend subroutine tag\nend program t\n",
        ["z"]
    };

    select_type_integer_modify_in_branch => {
        "program t\nclass(*), allocatable :: val\nallocate(integer :: val)\nval = 10\nselect type(val)\ntype is (integer)\nval = val + 5\nprint *, val\nend select\nend program t\n",
        ["15"]
    };

    select_rank_vector_minval => {
        "program t\ncall tag([8, 3, 12, 1])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, minval(x)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["1"]
    };

    select_rank_vector_maxval => {
        "program t\ncall tag([8, 3, 12, 1])\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(1)\nprint *, maxval(x)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["12"]
    };

    select_type_real_compare_in_branch => {
        "program t\nclass(*), allocatable :: val\nallocate(real :: val)\nval = 3.0\nselect type(val)\ntype is (real)\nif (val > 2.0) then\nprint *, 1\nelse\nprint *, 0\nend if\nend select\nend program t\n",
        ["1"]
    };

    select_rank_matrix_row_sum => {
        "program t\ncall tag(reshape([1,2,3,4],[2,2]))\ncontains\nsubroutine tag(x)\ninteger, intent(in) :: x(..)\nselect rank(x)\nrank(2)\nprint *, x(1,1) + x(1,2)\nrank default\nprint *, 0\nend select\nend subroutine tag\nend program t\n",
        ["3"]
    };
}

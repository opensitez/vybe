//! Extended WHERE / ELSEWHERE / MERGE coverage: masked assignment, array and scalar
//! forms, MERGE with logical masks, and nested conditions.
//! Distinct from `test_where_advanced.rs` (compile-only), `test_array_transforms.rs`
//! (basic MERGE scalars/arrays), and `test_arrays.rs` (basic WHERE compile).

fortran_cases! {
    // ── Masked assignment (no ELSEWHERE) ────────────────────────────

    where_masked_zero_negatives_only => {
        "program t\ninteger :: a(5)=[3,-1,7,-4,2]\nwhere (a<0)\na=0\nend where\nprint *,a(2)\nprint *,a(3)\nprint *,sum(a)\nend program t\n",
        ["0", "7", "12"]
    };
    where_masked_square_evens => {
        "program t\ninteger :: a(6)=[1,2,3,4,5,6]\nwhere (mod(a,2)==0)\na=a*a\nend where\nprint *,a(2)\nprint *,a(3)\nprint *,a(4)\nend program t\n",
        ["4", "3", "16"]
    };
    where_masked_double_gt_five => {
        "program t\ninteger :: a(5)=[2,6,3,8,1]\nwhere (a>5)\na=a*2\nend where\nprint *,a(2)\nprint *,a(3)\nprint *,a(4)\nend program t\n",
        ["12", "3", "16"]
    };
    where_masked_abs_copy => {
        "program t\ninteger :: a(4)=[-5,3,-2,7]\ninteger :: b(4)=0\nwhere (a<0)\nb=-a\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["5", "0", "2"]
    };
    where_masked_self_increment => {
        "program t\ninteger :: a(4)=[1,2,3,4]\nwhere (a>=3)\na=a+10\nend where\nprint *,a(3)\nprint *,a(4)\nprint *,sum(a)\nend program t\n",
        ["13", "14", "30"]
    };
    where_masked_mod_three_zero => {
        "program t\ninteger :: a(6)=[1,2,3,4,5,6]\nwhere (mod(a,3)==0)\na=0\nend where\nprint *,a(3)\nprint *,a(4)\nprint *,a(6)\nend program t\n",
        ["0", "4", "0"]
    };
    where_masked_real_neg_to_zero => {
        "program t\nreal :: a(4)=[1.5,-2.0,3.0,-0.5]\nwhere (a<0.0)\na=0.0\nend where\nprint *,a(2)\nprint *,a(3)\nend program t\n",
        ["0", "3"]
    };
    where_masked_logical_to_int => {
        "program t\ninteger :: a(4)=[-1,0,2,-3]\ninteger :: b(4)=0\nwhere (a>0)\nb=1\nend where\nprint *,b(1)\nprint *,b(3)\nprint *,sum(b)\nend program t\n",
        ["0", "1", "1"]
    };
    where_masked_2d_diagonal_double => {
        "program t\ninteger :: m(3,3)\ninteger :: i,j\nm=0\ndo i=1,3\nm(i,i)=i\nend do\nwhere (m>0)\nm=m*10\nend where\nprint *,m(1,1)\nprint *,m(2,2)\nprint *,m(1,2)\nend program t\n",
        ["10", "20", "0"]
    };
    where_masked_assign_from_other_array => {
        "program t\ninteger :: src(4)=[10,20,30,40]\ninteger :: dst(4)=[0,0,0,0]\ninteger :: mask(4)=[1,0,1,0]\nwhere (mask==1)\ndst=src\nend where\nprint *,dst(1)\nprint *,dst(2)\nprint *,dst(3)\nend program t\n",
        ["10", "0", "30"]
    };

    // ── Scalar WHERE mask ───────────────────────────────────────────

    where_scalar_mask_true_assigns_all => {
        "program t\ninteger :: a(3)=[1,2,3]\nwhere (.true.)\na=9\nend where\nprint *,a(1)\nprint *,a(3)\nprint *,sum(a)\nend program t\n",
        ["9", "9", "27"]
    };
    where_scalar_mask_false_leaves_unchanged => {
        "program t\ninteger :: a(3)=[4,5,6]\nwhere (.false.)\na=0\nend where\nprint *,a(1)\nprint *,sum(a)\nend program t\n",
        ["4", "15"]
    };
    where_scalar_rhs_assigns_constant => {
        "program t\ninteger :: a(4)=[1,2,3,4]\nwhere (a>2)\na=99\nend where\nprint *,a(2)\nprint *,a(3)\nprint *,a(4)\nend program t\n",
        ["2", "99", "99"]
    };
    where_scalar_lhs_array_rhs_scalar => {
        "program t\ninteger :: a(5)=[0,0,0,0,0]\ninteger :: b(5)=[1,2,3,4,5]\nwhere (b>3)\na=b\nend where\nprint *,a(4)\nprint *,a(5)\nprint *,sum(a)\nend program t\n",
        ["4", "5", "9"]
    };
    where_scalar_single_element_mask => {
        "program t\ninteger :: a(3)=[1,2,3]\nlogical :: m(3)=[.false.,.true.,.false.]\nwhere (m)\na=7\nend where\nprint *,a(1)\nprint *,a(2)\nprint *,a(3)\nend program t\n",
        ["1", "7", "3"]
    };

    // ── WHERE with ELSEWHERE ────────────────────────────────────────

    where_else_int_multiply_branch => {
        "program t\ninteger :: a(4)=[2,8,3,9]\ninteger :: b(4)\nwhere (a>5)\nb=a*10\nelsewhere\nb=a\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(4)\nend program t\n",
        ["2", "80", "90"]
    };
    where_else_real_clamp_negatives => {
        "program t\nreal :: x(4)=[-1.0,2.0,-3.0,4.0]\nreal :: y(4)\nwhere (x>=0.0)\ny=x\nelsewhere\ny=0.0\nend where\nprint *,y(1)\nprint *,y(2)\nprint *,y(3)\nend program t\n",
        ["0", "2", "0"]
    };
    where_else_neg_to_zero_pos_unchanged => {
        "program t\ninteger :: v(5)=[5,-2,8,-1,3]\nwhere (v<0)\nv=0\nelsewhere\nv=v\nend where\nprint *,v(2)\nprint *,v(3)\nprint *,sum(v)\nend program t\n",
        ["0", "8", "16"]
    };
    where_else_char_branch_labels => {
        "program t\ninteger :: s(3)=[1,5,12]\ncharacter(len=1) :: c(3)\nwhere (s<3)\nc=\"L\"\nelsewhere\nc=\"H\"\nend where\nprint *,c(1)\nprint *,c(2)\nprint *,c(3)\nend program t\n",
        ["L", "H", "H"]
    };
    where_else_2d_hi_lo_split => {
        "program t\ninteger :: m(2,2)=reshape([1,6,3,8],[2,2])\ninteger :: r(2,2)\nwhere (m>5)\nr=m*2\nelsewhere\nr=m\nend where\nprint *,r(1,1)\nprint *,r(1,2)\nprint *,r(2,2)\nend program t\n",
        ["1", "12", "16"]
    };
    where_else_increment_decrement => {
        "program t\ninteger :: a(4)=[1,4,7,10]\nwhere (mod(a,2)==0)\na=a+1\nelsewhere\na=a-1\nend where\nprint *,a(1)\nprint *,a(2)\nprint *,a(3)\nend program t\n",
        ["0", "5", "6"]
    };
    where_else_logical_int_encoding => {
        "program t\ninteger :: a(4)=[-1,0,2,5]\ninteger :: b(4)\nwhere (a>0)\nb=1\nelsewhere\nb=0\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["0", "0", "1"]
    };

    // ── Multiple ELSEWHERE clauses ──────────────────────────────────

    where_multi_else_sign_three_way => {
        "program t\ninteger :: v(3)=[-2,0,5]\ninteger :: s(3)\nwhere (v<0)\ns=-1\nelsewhere (v==0)\ns=0\nelsewhere\ns=1\nend where\nprint *,s(1)\nprint *,s(2)\nprint *,s(3)\nend program t\n",
        ["-1", "0", "1"]
    };
    where_multi_else_size_tiers => {
        "program t\ninteger :: a(4)=[3,15,50,200]\ninteger :: t(4)\nwhere (a<10)\nt=1\nelsewhere (a<100)\nt=2\nelsewhere\nt=3\nend where\nprint *,t(1)\nprint *,t(2)\nprint *,t(4)\nend program t\n",
        ["1", "2", "3"]
    };
    where_multi_else_temp_four_ranges => {
        "program t\nreal :: t(4)=[-10.0,-1.0,5.0,50.0]\ninteger :: c(4)\nwhere (t<0.0)\nc=0\nelsewhere (t<10.0)\nc=1\nelsewhere (t<40.0)\nc=2\nelsewhere\nc=3\nend where\nprint *,c(1)\nprint *,c(2)\nprint *,c(3)\nprint *,c(4)\nend program t\n",
        ["0", "1", "1", "3"]
    };
    where_multi_else_grade_buckets => {
        "program t\ninteger :: p(4)=[55,72,88,95]\ncharacter(len=1) :: g(4)\nwhere (p<60)\ng=\"F\"\nelsewhere (p<70)\ng=\"D\"\nelsewhere (p<80)\ng=\"C\"\nelsewhere\ng=\"A\"\nend where\nprint *,g(1)\nprint *,g(2)\nprint *,g(4)\nend program t\n",
        ["F", "C", "A"]
    };
    where_multi_else_real_quadrant => {
        "program t\nreal :: x(4)=[-3.0,2.0,-1.0,4.0]\ninteger :: q(4)\nwhere (x<0.0)\nq=1\nelsewhere (x<3.0)\nq=2\nelsewhere\nq=3\nend where\nprint *,q(1)\nprint *,q(2)\nprint *,q(4)\nend program t\n",
        ["1", "2", "3"]
    };

    // ── Nested WHERE ────────────────────────────────────────────────

    nested_where_even_odd_tiers => {
        "program t\ninteger :: a(4)=[2,3,4,5]\ninteger :: b(4)=0\nwhere (a>2)\nwhere (mod(a,2)==0)\nb=a*10\nelsewhere\nb=a\nend where\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["0", "3", "40"]
    };
    nested_where_magnitude_inner_else => {
        "program t\ninteger :: a(4)=[1,8,15,20]\ninteger :: b(4)=0\nwhere (a>5)\nwhere (a>12)\nb=2\nelsewhere\nb=1\nend where\nelsewhere\nb=0\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["0", "1", "2"]
    };
    nested_where_2d_positive_inner => {
        "program t\ninteger :: m(2,2)=reshape([1,-2,3,-4],[2,2])\ninteger :: r(2,2)=0\nwhere (m>0)\nwhere (m>2)\nr=m*2\nelsewhere\nr=m\nend where\nend where\nprint *,r(1,1)\nprint *,r(2,1)\nprint *,r(2,2)\nend program t\n",
        ["1", "6", "0"]
    };
    nested_where_outer_else_inner_match => {
        "program t\ninteger :: a(3)=[10,3,20]\ninteger :: b(3)=0\nwhere (a>5)\nwhere (a>15)\nb=2\nelsewhere\nb=1\nend where\nelsewhere\nb=-1\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["1", "-1", "2"]
    };
    nested_where_depth_two_increment => {
        "program t\ninteger :: a(4)=[1,6,11,16]\ninteger :: b(4)=0\nwhere (a>3)\nwhere (a>10)\nb=b+2\nelsewhere\nb=b+1\nend where\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["0", "1", "2"]
    };

    // ── MERGE with mask ─────────────────────────────────────────────

    merge_real_scalar_true_branch => {
        "program t\nreal :: x\nx=merge(3.5,7.5,.true.)\nprint *,x\nend program t\n",
        ["3.5"]
    };
    merge_real_scalar_false_branch => {
        "program t\nreal :: x\nx=merge(3.5,7.5,.false.)\nprint *,x\nend program t\n",
        ["7.5"]
    };
    merge_array_from_comparison_mask => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: b(4)=[10,20,30,40]\ninteger :: c(4)\nc=merge(a,b,a>2)\nprint *,c(1)\nprint *,c(2)\nprint *,c(4)\nend program t\n",
        ["10", "20", "40"]
    };
    merge_array_2d_mask => {
        "program t\ninteger :: a(2,2)=reshape([1,2,3,4],[2,2])\ninteger :: b(2,2)=reshape([9,8,7,6],[2,2])\ninteger :: c(2,2)\nc=merge(a,b,a<b)\nprint *,c(1,1)\nprint *,c(1,2)\nprint *,c(2,1)\nend program t\n",
        ["1", "8", "3"]
    };
    merge_logical_scalar_values => {
        "program t\nlogical :: x\nx=merge(.true.,.false.,.true.)\nprint *,x\nend program t\n",
        ["true"]
    };
    merge_char_array_pick => {
        "program t\ncharacter(len=1) :: a(3)=[\"A\",\"B\",\"C\"]\ncharacter(len=1) :: b(3)=[\"X\",\"Y\",\"Z\"]\nlogical :: m(3)=[.true.,.false.,.true.]\ncharacter(len=1) :: c(3)\nc=merge(a,b,m)\nprint *,c(1)\nprint *,c(2)\nprint *,c(3)\nend program t\n",
        ["A", "Y", "C"]
    };
    merge_mask_from_variable => {
        "program t\nlogical :: flag=.false.\ninteger :: x\nx=merge(100,200,flag)\nprint *,x\nend program t\n",
        ["200"]
    };
    merge_nested_twice => {
        "program t\ninteger :: x\nx=merge(merge(1,2,.true.),merge(3,4,.false.),.false.)\nprint *,x\nend program t\n",
        ["3"]
    };
    merge_in_where_body => {
        "program t\ninteger :: a(3)=[-1,2,-3]\ninteger :: b(3)\nwhere (a<0)\nb=merge(0,a,.true.)\nelsewhere\nb=a\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["0", "2", "0"]
    };
    merge_length5_alternating_mask => {
        "program t\ninteger :: a(5)=[1,1,1,1,1]\ninteger :: b(5)=[2,2,2,2,2]\nlogical :: m(5)=[.true.,.false.,.true.,.false.,.true.]\ninteger :: c(5)\nc=merge(a,b,m)\nprint *,c(1)\nprint *,c(2)\nprint *,c(5)\nprint *,sum(c)\nend program t\n",
        ["1", "2", "1", "7"]
    };

    // ── Combined WHERE + MERGE patterns ─────────────────────────────

    where_merge_equiv_positive_part => {
        "program t\ninteger :: a(4)=[-3,5,-1,8]\ninteger :: b(4)\ninteger :: i\ndo i=1,4\nb(i)=merge(a(i),0,a(i)>0)\nend do\nprint *,b(1)\nprint *,b(2)\nprint *,b(4)\nend program t\n",
        ["0", "5", "8"]
    };
    where_precomputed_logical_mask => {
        "program t\ninteger :: a(4)=[1,2,3,4]\nlogical :: m(4)=[.true.,.false.,.true.,.false.]\ninteger :: b(4)=0\nwhere (m)\nb=a*10\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,b(3)\nend program t\n",
        ["10", "0", "30"]
    };
    where_else_sum_of_branches => {
        "program t\ninteger :: a(4)=[1,6,2,9]\ninteger :: b(4)\nwhere (a>5)\nb=a*2\nelsewhere\nb=a+1\nend where\nprint *,b(1)\nprint *,b(2)\nprint *,sum(b)\nend program t\n",
        ["2", "12", "25"]
    };
    where_array_mask_all_false => {
        "program t\ninteger :: a(3)=[5,6,7]\ninteger :: b(3)=0\nwhere (a>10)\nb=1\nend where\nprint *,b(1)\nprint *,sum(b)\nend program t\n",
        ["0", "0"]
    };
    where_array_mask_all_true => {
        "program t\ninteger :: a(3)=[1,2,3]\ninteger :: b(3)=0\nwhere (a>0)\nb=a*2\nend where\nprint *,b(2)\nprint *,sum(b)\nend program t\n",
        ["4", "12"]
    };

    merge_mask_from_scalar_flag => {
        "program t\nlogical :: flag\ninteger :: x\nflag = .true.\nx = merge(10, 20, flag)\nprint *, x\nflag = .false.\nx = merge(10, 20, flag)\nprint *, x\nend program t\n",
        ["10", "20"]
    };

    merge_kind1_integer_arrays => {
        "program t\ninteger(kind=1) :: a(3)=[1_1, 2_1, 3_1]\ninteger(kind=1) :: b(3)=[4_1, 5_1, 6_1]\nlogical :: m(3)=[.true., .false., .true.]\ninteger(kind=1) :: c(3)\nc = merge(a, b, m)\nprint *, c(1)\nprint *, c(2)\nprint *, c(3)\nend program t\n",
        ["1", "5", "3"]
    };

    nested_where_with_merge_mask => {
        "program t\ninteger :: a(4)=[1,2,3,4]\ninteger :: b(4)=[4,3,2,1]\nlogical :: m(4)\nm = merge((/ .true., .false., .false., .true. /), (/ .false., .true., .true., .false. /), a > 2)\nwhere (m)\nb = a * 10\nelsewhere\nb = b - a\nend where\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nprint *, b(4)\nend program t\n",
        ["3", "1", "30", "-3"]
    };

    where_with_char_arrays => {
        "program t\ncharacter(len=2) :: src(3)=[\"aa\", \"bb\", \"cc\"]\ncharacter(len=2) :: dst(3)\nwhere (src /= \"bb\")\ndst = src\nelsewhere\ndst = \"--\"\nend where\nprint *, trim(dst(1))\nprint *, trim(dst(2))\nprint *, trim(dst(3))\nend program t\n",
        ["aa", "--", "cc"]
    };
}

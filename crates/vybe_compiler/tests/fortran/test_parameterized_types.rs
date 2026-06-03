use super::helpers::compile_ok;

// ── Length parameterized derived types (LEN parameter) ───────

#[test]
fn pdt_len_basic() {
    compile_ok(
        r#"
program test
    type :: FixedVec(n)
        integer, len :: n
        real :: data(n)
    end type FixedVec
    type(FixedVec(5)) :: v
    v%data = 0.0
    v%data(1) = 1.0
    print *, v%data(1)
end program test
"#,
    );
}

#[test]
fn pdt_len_string() {
    compile_ok(
        r#"
program test
    type :: BoundedStr(maxlen)
        integer, len :: maxlen
        character(len=maxlen) :: value
    end type BoundedStr
    type(BoundedStr(20)) :: s
    s%value = 'hello'
    print *, trim(s%value)
end program test
"#,
    );
}

#[test]
fn pdt_len_parameter() {
    compile_ok(
        r#"
program test
    type :: Matrix(m, n)
        integer, len :: m, n
        real :: data(m, n)
    end type Matrix
    type(Matrix(3,3)) :: mat
    mat%data = 0.0
    mat%data(2,2) = 99.0
    print *, mat%data(2,2)
end program test
"#,
    );
}

#[test]
fn pdt_len_in_subroutine() {
    compile_ok(
        r#"
program test
    type :: Vec(n)
        integer, len :: n
        real :: data(n)
    end type Vec
    type(Vec(4)) :: v
    v%data = [1.0, 2.0, 3.0, 4.0]
    call show_first(v)
contains
    subroutine show_first(v)
        type(Vec(*)), intent(in) :: v
        print *, v%data(1)
    end subroutine show_first
end program test
"#,
    );
}

#[test]
fn pdt_len_assumed_star() {
    compile_ok(
        r#"
program test
    type :: Buf(n)
        integer, len :: n
        integer :: items(n)
    end type Buf
    type(Buf(10)) :: b
    b%items = [(i, i=1,10)]
    call process(b)
contains
    subroutine process(x)
        type(Buf(*)), intent(in) :: x
        print *, x%items(1)
        print *, size(x%items)
    end subroutine process
end program test
"#,
    );
}

// ── Kind parameterized derived types (KIND parameter) ─────────

#[test]
fn pdt_kind_basic() {
    compile_ok(
        r#"
program test
    type :: TypedNum(k)
        integer, kind :: k
        real(k) :: value
    end type TypedNum
    type(TypedNum(4)) :: f
    type(TypedNum(8)) :: d
    f%value = 3.14_4
    d%value = 3.14159265358979_8
    print *, f%value
    print *, d%value
end program test
"#,
    );
}

#[test]
fn pdt_kind_complex() {
    compile_ok(
        r#"
program test
    type :: TypedComplex(k)
        integer, kind :: k
        complex(k) :: value
    end type TypedComplex
    type(TypedComplex(4)) :: c
    c%value = (1.0_4, 2.0_4)
    print *, real(c%value)
end program test
"#,
    );
}

#[test]
fn pdt_kind_integer() {
    compile_ok(
        r#"
program test
    type :: TypedInt(k)
        integer, kind :: k
        integer(k) :: value
    end type TypedInt
    type(TypedInt(8)) :: big
    big%value = 1000000000_8
    print *, big%value
end program test
"#,
    );
}

// ── Both LEN and KIND parameters ──────────────────────────────

#[test]
fn pdt_len_and_kind() {
    compile_ok(
        r#"
program test
    type :: Precision(k, n)
        integer, kind :: k
        integer, len :: n
        real(k) :: data(n)
    end type Precision
    type(Precision(8, 5)) :: p
    p%data = 1.0_8
    p%data(3) = 3.14159265_8
    print *, p%data(3)
end program test
"#,
    );
}

// ── Default parameter values ───────────────────────────────────

#[test]
fn pdt_default_len() {
    compile_ok(
        r#"
program test
    type :: DefaultVec(n)
        integer, len :: n = 10
        real :: data(n)
    end type DefaultVec
    type(DefaultVec()) :: v
    v%data = 0.0
    print *, size(v%data)
end program test
"#,
    );
}

#[test]
fn pdt_default_kind() {
    compile_ok(
        r#"
program test
    type :: DefaultReal(k)
        integer, kind :: k = 4
        real(k) :: x
    end type DefaultReal
    type(DefaultReal()) :: r
    r%x = 1.0
    print *, r%x
end program test
"#,
    );
}

// ── PDT arrays ────────────────────────────────────────────────

#[test]
fn pdt_array_of_pdt() {
    compile_ok(
        r#"
program test
    type :: Pair(k)
        integer, kind :: k
        real(k) :: x, y
    end type Pair
    type(Pair(4)) :: pairs(3)
    integer :: i
    do i = 1, 3
        pairs(i)%x = real(i, 4)
        pairs(i)%y = real(i, 4) * 2.0_4
    end do
    print *, pairs(2)%x
end program test
"#,
    );
}

// ── PDT allocatable ───────────────────────────────────────────

#[test]
fn pdt_deferred_len() {
    compile_ok(
        r#"
program test
    type :: DynVec(n)
        integer, len :: n
        real :: data(n)
    end type DynVec
    type(DynVec(:)), allocatable :: v
    allocate(DynVec(10) :: v)
    v%data = 0.0
    v%data(5) = 42.0
    print *, v%data(5)
    deallocate(v)
end program test
"#,
    );
}

#[test]
fn pdt_allocatable_component() {
    compile_ok(
        r#"
program test
    type :: DynMat(m, n)
        integer, len :: m, n
        real, allocatable :: data(:,:)
    end type DynMat
    type(DynMat(3,4)) :: mat
    allocate(mat%data(mat%m, mat%n))
    mat%data = 0.0
    mat%data(2,3) = 7.0
    print *, mat%data(2,3)
    deallocate(mat%data)
end program test
"#,
    );
}

// ── PDT in modules ────────────────────────────────────────────

#[test]
fn pdt_in_module() {
    compile_ok(
        r#"
module pdt_mod
    implicit none
    type :: Tensor(rank, k)
        integer, len  :: rank
        integer, kind :: k
        real(k) :: components(rank)
    end type Tensor
contains
    subroutine zero(t)
        type(Tensor(*,*)), intent(inout) :: t
        t%components = 0
    end subroutine zero
end module pdt_mod

program test
    use pdt_mod
    type(Tensor(3,4)) :: v
    call zero(v)
    print *, v%components(1)
end program test
"#,
    );
}

// ── PDT inquiry functions ─────────────────────────────────────

#[test]
fn pdt_len_inquiry() {
    compile_ok(
        r#"
program test
    type :: Str(n)
        integer, len :: n
        character(n) :: s
    end type Str
    type(Str(15)) :: t
    t%s = 'hello'
    print *, t%n
end program test
"#,
    );
}

#[test]
fn pdt_kind_inquiry() {
    compile_ok(
        r#"
program test
    type :: Num(k)
        integer, kind :: k
        real(k) :: v
    end type Num
    type(Num(8)) :: x
    x%v = 1.0_8
    print *, x%k
end program test
"#,
    );
}

use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    allocate_statement_01,
    "program p
integer, allocatable :: a(:)
allocate(a(3))
end program p
"
);
c!(
    allocate_statement_02,
    "program p
real, allocatable :: a(:,:)
allocate(a(2,2))
end program p
"
);
c!(
    allocate_statement_03,
    "program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
end program p
"
);
c!(
    allocate_statement_04,
    "program p
integer, pointer :: p(:)
allocate(p(3))
end program p
"
);
c!(
    allocate_statement_05,
    "program p
integer, allocatable :: x
allocate(x)
end program p
"
);
c!(
    allocate_statement_06,
    "program p
complex, allocatable :: z(:)
allocate(z(2))
end program p
"
);
c!(
    allocate_statement_07,
    "program p
logical, allocatable :: l(:)
allocate(l(2))
end program p
"
);
c!(
    allocate_statement_08,
    "program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,2))
end program p
"
);
c!(
    allocate_statement_09,
    "program p
type t
 integer :: x
end type t
type(t), allocatable :: a(:)
allocate(a(2))
end program p
"
);
c!(
    allocate_statement_10,
    "program p
class(*), allocatable :: x
allocate(integer :: x)
end program p
"
);

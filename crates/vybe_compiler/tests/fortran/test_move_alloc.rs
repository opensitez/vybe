use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(move_alloc_01,"program p
integer, allocatable :: a(:), b(:)
allocate(a(3))
call move_alloc(a,b)
end program p
");
c!(move_alloc_02,"program p
integer, allocatable :: a,b
allocate(a)
call move_alloc(a,b)
end program p
");
c!(move_alloc_03,"program p
real, allocatable :: a(:,:), b(:,:)
allocate(a(2,2))
call move_alloc(a,b)
end program p
");
c!(move_alloc_04,"program p
character(len=:), allocatable :: a,b
allocate(character(len=3) :: a)
call move_alloc(a,b)
end program p
");
c!(move_alloc_05,"program p
logical, allocatable :: a(:), b(:)
allocate(a(2))
call move_alloc(a,b)
end program p
");
c!(move_alloc_06,"program p
complex, allocatable :: a(:), b(:)
allocate(a(2))
call move_alloc(a,b)
end program p
");
c!(move_alloc_07,"type t
 integer :: x
end type t
program p
type(t), allocatable :: a,b
allocate(a)
call move_alloc(a,b)
end program p
");
c!(move_alloc_08,"program p
integer, allocatable :: a(:,:,:), b(:,:,:)
allocate(a(2,2,2))
call move_alloc(a,b)
end program p
");
c!(move_alloc_09,"program p
integer, allocatable :: a(:)
allocate(a(0))
call move_alloc(a,a)
end program p
");
c!(move_alloc_10,"program p
class(*), allocatable :: a,b
allocate(integer :: a)
call move_alloc(a,b)
end program p
");
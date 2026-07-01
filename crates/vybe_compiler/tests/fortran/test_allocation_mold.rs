use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(allocation_mold_01,"program p
integer, allocatable :: a(:), b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
");
c!(allocation_mold_02,"program p
integer, allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
");
c!(allocation_mold_03,"program p
real, allocatable :: a(:,:), b(:,:)
allocate(b(2,2))
allocate(a, mold=b)
end program p
");
c!(allocation_mold_04,"program p
character(len=:), allocatable :: s, t
allocate(character(len=4) :: t)
allocate(s, mold=t)
end program p
");
c!(allocation_mold_05,"program p
logical, allocatable :: a(:), b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
");
c!(allocation_mold_06,"program p
complex, allocatable :: a(:), b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
");
c!(allocation_mold_07,"type t
 integer :: x
end type t
program p
type(t), allocatable :: a,b
allocate(b)
allocate(a, mold=b)
end program p
");
c!(allocation_mold_08,"program p
class(*), allocatable :: x, y
allocate(integer :: y)
allocate(x, mold=y)
end program p
");
c!(allocation_mold_09,"program p
integer, pointer :: p(:)
integer, allocatable :: a(:)
allocate(p(3))
allocate(a, mold=p)
end program p
");
c!(allocation_mold_10,"program p
real, allocatable :: a(:,:,:), b(:,:,:)
allocate(b(2,2,2))
allocate(a, mold=b)
end program p
");
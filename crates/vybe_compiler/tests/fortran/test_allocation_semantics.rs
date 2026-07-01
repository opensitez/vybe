use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(as_01,"program p
integer, allocatable :: a(:)
allocate(a(3))
print *, size(a)
end program p
");
c!(as_02,"program p
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
print *, a
end program p
");
c!(as_03,"program p
integer, allocatable :: a(:),b(:)
allocate(b(2))
allocate(a, mold=b)
end program p
");
c!(as_04,"program p
integer, allocatable :: a(:),b(:)
allocate(a(3))
call move_alloc(a,b)
end program p
");
c!(as_05,"program p
integer, pointer :: p(:)
allocate(p(3))
end program p
");
c!(as_06,"program p
real, allocatable :: a(:,:)
allocate(a(2,3))
end program p
");
c!(as_07,"program p
complex, allocatable :: a(:)
allocate(a(4))
end program p
");
c!(as_08,"program p
logical, allocatable :: a(:)
allocate(a(5))
end program p
");
c!(as_09,"program p
character(len=:), allocatable :: s
allocate(character(len=5) :: s)
end program p
");
c!(as_10,"program p
integer, allocatable :: x
allocate(x, source=5)
end program p
");
c!(as_11,"program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3), stat=st)
end program p
");
c!(as_12,"program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3))
deallocate(a, stat=st)
end program p
");
c!(as_13,"type :: t
integer, allocatable :: a(:)
end type t
program p
type(t) :: x
allocate(x%a(2))
end program p
");
c!(as_14,"type :: t
character(len=:), allocatable :: s
end type t
program p
type(t) :: x
allocate(character(len=3) :: x%s)
end program p
");
c!(as_15,"program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,2))
end program p
");
c!(as_16,"program p
integer, allocatable :: a(:,:,:,:)
allocate(a(2,2,2,2))
end program p
");
c!(as_17,"program p
integer, allocatable :: a(:)
allocate(a(0))
print *, size(a)
end program p
");
c!(as_18,"program p
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a)
end program p
");
c!(as_19,"program p
integer, allocatable :: a(:)
allocate(a(3))
a = [1,2,3]
end program p
");
c!(as_20,"program p
class(*), allocatable :: x
allocate(integer :: x)
end program p
");
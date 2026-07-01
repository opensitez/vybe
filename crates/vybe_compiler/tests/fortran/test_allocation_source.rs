use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(allocation_source_01,"program p
integer, allocatable :: a(:)
allocate(a(3), source=[1,2,3])
end program p
");
c!(allocation_source_02,"program p
integer, allocatable :: x
allocate(x, source=5)
end program p
");
c!(allocation_source_03,"program p
real, allocatable :: a(:,:)
allocate(a(2,2), source=reshape([1.0,2.0,3.0,4.0],[2,2]))
end program p
");
c!(allocation_source_04,"program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s, source='abc')
end program p
");
c!(allocation_source_05,"program p
character(len=:), allocatable :: a(:)
allocate(character(len=2) :: a(2), source=['aa','bb'])
end program p
");
c!(allocation_source_06,"type t
 integer :: x
end type t
program p
type(t), allocatable :: v
allocate(v, source=t(1))
end program p
");
c!(allocation_source_07,"program p
logical, allocatable :: l(:)
allocate(l(2), source=[.true.,.false.])
end program p
");
c!(allocation_source_08,"program p
complex, allocatable :: z(:)
allocate(z(2), source=[(1.0,2.0),(3.0,4.0)])
end program p
");
c!(allocation_source_09,"program p
class(*), allocatable :: x
allocate(integer :: x, source=1)
end program p
");
c!(allocation_source_10,"program p
integer, allocatable :: a(:,:,:)
allocate(a(2,2,1), source=reshape([1,2,3,4],[2,2,1]))
end program p
");
use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(deallocate_statement_01,"program p
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a)
end program p
");
c!(deallocate_statement_02,"program p
real, allocatable :: a(:,:)
allocate(a(2,2))
deallocate(a)
end program p
");
c!(deallocate_statement_03,"program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
deallocate(s)
end program p
");
c!(deallocate_statement_04,"program p
integer, allocatable :: x
allocate(x)
deallocate(x)
end program p
");
c!(deallocate_statement_05,"program p
integer, allocatable :: a(:), b(:)
allocate(a(2),b(2))
deallocate(a,b)
end program p
");
c!(deallocate_statement_06,"program p
integer, pointer :: p(:)
allocate(p(3))
deallocate(p)
end program p
");
c!(deallocate_statement_07,"program p
class(*), allocatable :: x
allocate(integer :: x)
deallocate(x)
end program p
");
c!(deallocate_statement_08,"program p
type t
 integer :: x
end type t
type(t), allocatable :: a(:)
allocate(a(2))
deallocate(a)
end program p
");
c!(deallocate_statement_09,"program p
logical, allocatable :: l(:)
allocate(l(2))
deallocate(l)
end program p
");
c!(deallocate_statement_10,"program p
complex, allocatable :: z(:)
allocate(z(2))
deallocate(z)
end program p
");
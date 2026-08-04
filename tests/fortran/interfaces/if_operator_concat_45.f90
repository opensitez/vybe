! vybe-test: fortran/interfaces/if_operator_concat_45
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface operator(//)
module procedure concat_i
end interface
contains
character(len=20) function concat_i(a, b)
integer, intent(in) :: a, b
character(len=20) :: sa, sb
write(sa, '(I0)') a
write(sb, '(I0)') b
concat_i = trim(sa) // trim(sb)
end function concat_i
end module m

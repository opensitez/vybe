! vybe-test: fortran/oop/oop_inherit_chain_03
! origin: languages/fortran/tests/fortran/test_oop.rs
program t
type::a
integer::x
end type a
type,extends(a)::b
integer::y
end type b
type,extends(b)::c
integer::z
end type c
type(c) :: obj
obj%x = 1
obj%y = 2
obj%z = 3
if (obj%x + obj%y + obj%z /= 6) then
    print *, "FAIL: want [6] got [", obj%x + obj%y + obj%z, "]"
    stop 1
end if
if (obj%b%a%x /= 1) then
    print *, "FAIL: want [1] got [", obj%b%a%x, "]"
    stop 1
end if
end program t

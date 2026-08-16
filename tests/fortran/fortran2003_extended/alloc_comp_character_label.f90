! vybe-test: fortran/fortran2003_extended/alloc_comp_character_label
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program driver
type :: Tag
character(len=6), allocatable :: name
end type Tag
type(Tag) :: t
allocate(t%name)
t%name = 'f2003'
if (trim(trim(t%name)) /= "f2003") then
    print *, "FAIL: want [f2003] got [", trim(t%name), "]"
    stop 1
end if
end program driver
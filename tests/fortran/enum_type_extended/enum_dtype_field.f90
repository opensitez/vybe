! vybe-test: fortran/enum_type_extended/enum_dtype_field
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: IDLE = 0, RUN = 1, DONE = 2
end enum
type :: Task
integer :: state
end type Task
type(Task) :: t
t%state = RUN
if ((t%state) /= 1) then
    print *, "FAIL: want [1] got [", t%state, "]"
    stop 1
end if
end program t

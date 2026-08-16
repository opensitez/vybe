! vybe-test: fortran/sql/sql_delete_05
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program sql_delete_05
use iso_c_binding
implicit none
interface
  function sqlite3_open(filename, ppDb) bind(c, name="sqlite3_open") result(rc)
    import :: c_char, c_ptr, c_int
    character(kind=c_char), dimension(*) :: filename
    type(c_ptr) :: ppDb
    integer(c_int) :: rc
  end function
  function sqlite3_exec(db, sql, cb, arg, errmsg) bind(c, name="sqlite3_exec") result(rc)
    import :: c_char, c_ptr, c_int, c_funptr
    type(c_ptr), value :: db
    character(kind=c_char), dimension(*) :: sql
    type(c_funptr), value :: cb
    type(c_ptr), value :: arg
    type(c_ptr) :: errmsg
    integer(c_int) :: rc
  end function
  function sqlite3_close(db) bind(c, name="sqlite3_close") result(rc)
    import :: c_ptr, c_int
    type(c_ptr), value :: db
    integer(c_int) :: rc
  end function
end interface
type(c_ptr) :: db, errmsg
integer(c_int) :: rc
db = c_null_ptr
errmsg = c_null_ptr
rc = sqlite3_open(c_char_"sql_delete_05.db"//c_null_char, db)
if (rc /= 0) then
    print *, "FAIL: want [0] got [", rc, "]"
    stop 1
end if
rc = sqlite3_exec(db, c_char_"drop table if exists t"//c_null_char, c_null_funptr, c_null_ptr, errmsg)
if (rc /= 0) then
    print *, "FAIL: want [0] got [", rc, "]"
    stop 1
end if
rc = sqlite3_exec(db, c_char_"create table t(id int)"//c_null_char, c_null_funptr, c_null_ptr, errmsg)
if (rc /= 0) then
    print *, "FAIL: want [0] got [", rc, "]"
    stop 1
end if
rc = sqlite3_exec(db, c_char_"insert into t values(1)"//c_null_char, c_null_funptr, c_null_ptr, errmsg)
if (rc /= 0) then
    print *, "FAIL: want [0] got [", rc, "]"
    stop 1
end if
rc = sqlite3_exec(db, c_char_"delete from t where id = 1"//c_null_char, c_null_funptr, c_null_ptr, errmsg)
if (rc /= 0) then
    print *, "FAIL: want [0] got [", rc, "]"
    stop 1
end if
rc = sqlite3_close(db)
if (rc /= 0) then
    print *, "FAIL: want [0] got [", rc, "]"
    stop 1
end if
end program sql_delete_05

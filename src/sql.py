from RFDData import rednmx_query

def get_data():
    query = r'''
    SELECT 
        v.perscode AS Employee_ID,
        v.LNAME AS Last_Name,
        v.FNAME AS First_Name,
        v.schdshiftnamedescr AS Shift_Assignment,
        v.schdrankdescr AS Rank,
        v.shiftlength as Shift_Hours,
        v.datetimestart,
        DATEDIFF(DAY, p.dateappt, GETDATE()) / 365.25 AS YearsOfService,
        CASE WHEN v.schdtypeid IN (4, 7, 86,88) THEN 1 ELSE 0 END AS Missed_Flag
    FROM VWSCHDHIST v
    left join pers p on v.persid = p.persid
    WHERE v.schdrankid NOT IN (2, 9, 11, 14, 16)
        AND v.datetimestart BETWEEN DATEFROMPARTS(YEAR(GETDATE()), 1, 1) AND GETDATE()
        AND p.persstatid = 2
    ORDER BY v.perscode, v.datetimestart DESC;
    '''

    df = rednmx_query(query)
    return df 
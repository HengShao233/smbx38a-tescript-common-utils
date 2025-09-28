' 曲线 / 数学工具集
'                       -- 20231206 小豆

' ==================================================
' 使用简述:
' 计算前先 Call CUSetP0(...) ~ CUSetP3(...) 设置好各个控制点的位置
' 而后运行 CUCalcBezier__(t) 函数获得计算结果, 其中 t 取 0~1

' ==================================================
' 内建变量

' 常量
Dim MAX_FLOAT As Double = 3.402823466 * 10^38
Dim MIN_FLOAT As Double = 1.40129846432482 * 10^-45
Dim MAX_SAFE_DOUBLE As Double = 562949953421312

Export Script CUMath_MAX_SAFE_INTEGER(Return Double)
    Return MAX_SAFE_DOUBLE
End Script

Export Script CUMath_MAX_SAFE_FLOAT(Return Double)
    Return MAX_FLOAT
End Script

Export Script CUMath_MIN_SAFE_FLOAT(Return Double)
    Return MIN_FLOAT
End Script

' 顶点数据
Dim libCurve_p0_x As Double = 0
Dim libCurve_p0_y As Double = 0
Dim libCurve_p1_x As Double = 0
Dim libCurve_p1_y As Double = 0
Dim libCurve_p2_x As Double = 0
Dim libCurve_p2_y As Double = 0
Dim libCurve_p3_x As Double = 0
Dim libCurve_p3_y As Double = 0

' 长度数据
Dim libCurve_bezier_lenght3 As Double = -1
Dim libCurve_bezier_lenght2 As Double = -1
Dim libCurve_bezier_lenght1 As Double = -1

' 临时变量
Dim libCurve_tempA As Double = 0           ' 临时变量 A
Dim libCurve_tempB As Double = 0           ' 临时变量 B
Dim libCurve_tempC As Byte = 0             ' 临时变量 C
Dim libCurve_tempD As Double = 0           ' 临时变量 D
Dim libCurve_tempE As Double = 0           ' 临时变量 E
Dim libCurve_tempF As Double = 0           ' 临时变量 F
Dim libCurve_tempG As Double = 0           ' 临时变量 G
Dim libCurve_t As Double = 0               ' t
Dim libCurve_t2 As Double = 0              ' t^2
Dim libCurve_t3 As Double = 0              ' t^3
Dim libCurve_n_t As Double = 1             ' 1 - t
Dim libCurve_n_t2 As Double = 1            ' (1 - t)^2
Dim libCurve_n_t3 As Double = 1            ' (1 - t)^3
Dim libCurve_timeStampSplitCnt As Long = 1 ' 时间戳分割
Dim libCurve_timeStampStart As Double = 0  ' 时间戳
Dim libCurve_timeStampEnd As Double = 0    ' 时间戳

Dim libCurve_retX As Double = 0           ' 返回值 X
Dim libCurve_retY As Double = 0             ' 返回值 Y

' ==================================================
' 设置各个控制点的参数

' 设置 0 点
Export Script CUSetP0(x As Double, y As Double, Return Integer)
    If Abs(x - libCurve_p0_x) >= 0.01 Or Abs(y - libCurve_p0_y) >= 0.01 Then
        libCurve_bezier_lenght1 = -1
        libCurve_bezier_lenght2 = -1
        libCurve_bezier_lenght3 = -1
        libCurve_p0_x = x
        libCurve_p0_y = y
        Return 0
    End If
    Return 1
End Script

' 设置 1 点
Export Script CUSetP1(x As Double, y As Double, Return Integer)
    If Abs(x - libCurve_p1_x) >= 0.01 Or Abs(y - libCurve_p1_y) >= 0.01 Then
        libCurve_bezier_lenght1 = -1
        libCurve_bezier_lenght2 = -1
        libCurve_bezier_lenght3 = -1
        libCurve_p1_x = x
        libCurve_p1_y = y
        Return 0
    End If
    Return 1
End Script

' 设置 2 点
Export Script CUSetP2(x As Double, y As Double, Return Integer)
    If Abs(x - libCurve_p2_x) >= 0.01 Or Abs(y - libCurve_p2_y) >= 0.01 Then
        libCurve_bezier_lenght2 = -1
        libCurve_bezier_lenght3 = -1
        libCurve_p2_x = x
        libCurve_p2_y = y
        Return 0
    End If
    Return 1
End Script

' 设置 3 点
Export Script CUSetP3(x As Double, y As Double, Return Integer)
    If Abs(x - libCurve_p3_x) >= 0.01 Or Abs(y - libCurve_p3_y) >= 0.01 Then
        libCurve_bezier_lenght3 = -1
        libCurve_p3_x = x
        libCurve_p3_y = y
        Return 0
    End If
    Return 1
End Script

' ==================================================
' 计算方法

' ------------------------------------
' ----------------------- 内建方法 BEGIN

Script CUInner_SetT(t As Double, Return Integer)
    If Abs(t - libCurve_t) >= 0.000000001 Then
        libCurve_t = t
        libCurve_t2 = t * t
        libCurve_t3 = libCurve_t2 * t

        libCurve_n_t = 1 - t
        libCurve_n_t2 = libCurve_n_t * libCurve_n_t
        libCurve_n_t3 = libCurve_n_t2 * libCurve_n_t
    End If
    Return 0
End Script

Script CUInner_CalcBezier3(p0 As Double, p1 As Double, p2 As Double, p3 As Double, Return Double)
    Return libCurve_n_t3 * p0 + 3 * libCurve_n_t2 * libCurve_t * p1 + 3 * libCurve_n_t * libCurve_t2 * p2 + libCurve_t3 * p3
End Script

Script CUInner_CalcBezier2(p0 As Double, p1 As Double, p2 As Double, Return Double)
    Return libCurve_n_t2 * p0 + 2 * libCurve_n_t * libCurve_t * p1 + libCurve_t2 * p2
End Script

Script CUInner_CalcBezier1(p0 As Double, p1 As Double, Return Double)
    Return libCurve_n_t2 * p0 + libCurve_t2 * p1
End Script

Script CUInner_CalcDistancePP(x1 As Double, y1 As Double, x2 As Double, y2 As Double, Return Double)
    x1 = x1 - x2
    y1 = y1 - y2
    Return Sqr(x1 * x1 + y1 * y1)
End Script

' ----------------------- 内建方法 END
' ------------------------------------

' ------------------------------------
' ----------------------- 计算贝塞尔曲线结果坐标 BEGIN

' 计算三次曲线 x 坐标
Export Script CUCalcBezier3X(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier3(libCurve_p0_x, libCurve_p1_x, libCurve_p2_x, libCurve_p3_x)
End Script

' 计算三次曲线 y 坐标
Export Script CUCalcBezier3Y(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier3(libCurve_p0_y, libCurve_p1_y, libCurve_p2_y, libCurve_p3_y)
End Script

' 计算二次曲线 x 坐标
Export Script CUCalcBezier2X(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier2(libCurve_p0_x, libCurve_p1_x, libCurve_p2_x)
End Script

' 计算二次曲线 y 坐标
Export Script CUCalcBezier2Y(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier2(libCurve_p0_y, libCurve_p1_y, libCurve_p2_y)
End Script

' 计算一次曲线 x 坐标
Export Script CUCalcBezier1X(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier1(libCurve_p0_x, libCurve_p1_x)
End Script

' 计算一次曲线 y 坐标
Export Script CUCalcBezier1Y(t As Double, Return Double)
    Call CUInner_SetT(t)
    Return CUInner_CalcBezier1(libCurve_p0_y, libCurve_p1_y)
End Script

' ----------------------- 计算贝塞尔曲线结果坐标 END
' ------------------------------------

' ------------------------------------
' ----------------------- 计算贝塞尔曲线长度 BEGIN

' 近似计算三阶贝塞尔曲线长度
Export Script CUCalcBezier3Len(Return Double)
    If libCurve_bezier_lenght3 > 0 Then
        Return libCurve_bezier_lenght3
    End If

    libCurve_tempB = 0
    libCurve_tempD = libCurve_p0_x
    libCurve_tempE = libCurve_p0_y
    For libCurve_tempC = 1 To 25
        libCurve_tempA = libCurve_tempC / 25

        libCurve_tempF = CUCalcBezier3X(libCurve_tempA)
        libCurve_tempG = CUCalcBezier3Y(libCurve_tempA)
        libCurve_tempB += CUInner_CalcDistancePP(libCurve_tempD, libCurve_tempE, libCurve_tempF, libCurve_tempG)
        libCurve_tempD = libCurve_tempF
        libCurve_tempE = libCurve_tempG
    Next
    libCurve_bezier_lenght3 = libCurve_tempB
    Return libCurve_bezier_lenght3
End Script

' 近似计算二阶贝塞尔曲线长度
Export Script CUCalcBezier2Len(Return Double)
    If libCurve_bezier_lenght2 > 0 Then
        Return libCurve_bezier_lenght2
    End If

    libCurve_tempB = 0
    libCurve_tempD = libCurve_p0_x
    libCurve_tempE = libCurve_p0_y
    For libCurve_tempC = 1 To 25
        libCurve_tempA = libCurve_tempC / 25

        libCurve_tempF = CUCalcBezier2X(libCurve_tempA)
        libCurve_tempG = CUCalcBezier2Y(libCurve_tempA)
        libCurve_tempB += CUInner_CalcDistancePP(libCurve_tempD, libCurve_tempE, libCurve_tempF, libCurve_tempG)
        libCurve_tempD = libCurve_tempF
        libCurve_tempE = libCurve_tempG
    Next
    libCurve_bezier_lenght2 = libCurve_tempB
    Return libCurve_bezier_lenght2
End Script

' 计算一阶贝塞尔曲线长度
Export Script CUCalcBezier1Len(Return Double)
    If libCurve_bezier_lenght1 > 0 Then
        Return libCurve_bezier_lenght1
    End If

    libCurve_bezier_lenght1 = CUInner_CalcDistancePP(libCurve_p0_x, libCurve_p0_y, libCurve_p1_x, libCurve_p1_y)
    Return libCurve_bezier_lenght1
End Script

' ----------------------- 计算贝塞尔曲线长度 END
' ------------------------------------

' ------------------------------------
' ----------------------- 应用曲线 BEGIN

' 应用曲线到 bitmap 位置坐标
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpPosBezier3(id As Long, t As Double, Return Double)
    Bitmap(id).destx = CUCalcBezier3X(t)
    Bitmap(id).desty = CUCalcBezier3Y(t)
    Return t
End Script

' 应用曲线到 bitmap 位置坐标
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpPosBezier2(id As Long, t As Double, Return Double)
    Bitmap(id).destx = CUCalcBezier2X(t)
    Bitmap(id).desty = CUCalcBezier2Y(t)
    Return t
End Script

' 应用曲线到 bitmap 位置坐标
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpPosBezier1(id As Long, t As Double, Return Double)
    Bitmap(id).destx = CUCalcBezier1X(t)
    Bitmap(id).desty = CUCalcBezier1Y(t)
    Return t
End Script

' 应用曲线到 bitmap 缩放
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpScaleBezier3(id As Long, t As Double, Return Double)
    Bitmap(id).scalex = CUCalcBezier3X(t)
    Bitmap(id).scaley = CUCalcBezier3Y(t)
    Return t
End Script

' 应用曲线到 bitmap 缩放
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpScaleBezier2(id As Long, t As Double, Return Double)
    Bitmap(id).scalex = CUCalcBezier2X(t)
    Bitmap(id).scaley = CUCalcBezier2Y(t)
    Return t
End Script

' 应用曲线到 bitmap 缩放
' params id bitmap id
' params t 曲线参数
Export Script CUSetBmpScaleBezier1(id As Long, t As Double, Return Double)
    Bitmap(id).scalex = CUCalcBezier1X(t)
    Bitmap(id).scaley = CUCalcBezier1Y(t)
    Return t
End Script

' ----------------------- 应用曲线 END
' ------------------------------------

' ------------------------------------
' ----------------------- 时间工具集 BEGIN

' 设置时间戳端点
' @params startTimeStamp 时间戳起点
' @params endTimeStamp 时间戳终点
Export Script CUTimeSetStamp(startTimeStamp As Double, endTimeStamp As Double, Return Double)
    libCurve_timeStampEnd = endTimeStamp
    libCurve_timeStampStart = startTimeStamp
    Return endTimeStamp - startTimeStamp
End Script

' 设置时间戳分割数
' @params timeStampSplit 时间戳分割数
Export Script CUTimeSetSplit(timeStampSplit As Long, Return Long)
    If timeStampSplit <= 0 Then
        libCurve_timeStampSplitCnt = 1
    Else
        libCurve_timeStampSplitCnt = timeStampSplit
    End If
    Return timeStampSplit
End Script

' 计算 t 参数
' @params timeStamp 当前时间戳
' @return t 参数
Export Script CUTimeCalcT(timeStamp As Double, Return Double)
    If Abs(libCurve_timeStampEnd - libCurve_timeStampStart) < 0.000000001 Then
        Return 1
    End If

    timeStamp = (timeStamp - libCurve_timeStampStart) / (libCurve_timeStampEnd - libCurve_timeStampStart)
    If timeStamp < 0 Then
        Return 0
    ElseIf timeStamp > 1 Then
        Return 1
    End If

    Return timeStamp
End Script

' 计算已分段的 t 参数
' @params t 当前全局时间参数
' @return 返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitT(t As Double, Return Double)
    If t >= 1 Then
        Return 1
    ElseIf t <= 0 Then
        Return 0
    End If
    t = t * libCurve_timeStampSplitCnt
    Return t - Int(t)
End Script

' 计算当前所处段号
' @params t 当前全局时间参数
' @return 当前所处段号
Export Script CUTimeCalcSplitIdx(t As Double, Return Long)
    Return Int(t * libCurve_timeStampSplitCnt)
End Script

' 计算已分段 t 参数 (t∈[0, 1])
' @params t 当前全局时间参数
' @params curSplitIdx 当前段号
' @return 若当前 t 不属于当前段号则返回 -1, 否则返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitTByIdx(t As Double, curSplitIdx As Long, Return Double)
    If curSplitIdx <= 0 Or curSplitIdx > libCurve_timeStampSplitCnt - 1 Then
        Return -1
    End If

    t = t * libCurve_timeStampSplitCnt
    If t < curSplitIdx Or t - curSplitIdx > 1.000000001 Then
        Return -1
    End If

    Return t - curSplitIdx
End Script

' 计算已分段的 t 参数
' @params timeStamp 当前时间戳
' @return 返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitTByStamp(timeStamp As Double, Return Double)
    timeStamp = CUTimeCalcT(timeStamp)
    timeStamp = timeStamp * libCurve_timeStampSplitCnt
    Return timeStamp - Int(timeStamp)
End Script

' 计算当前所处段号
' @params timeStamp 当前时间戳
' @return 当前所处段号
Export Script CUTimeCalcSplitIdxByStamp(timeStamp As Double, Return Long)
    timeStamp = CUTimeCalcT(timeStamp)
    Return Int(timeStamp * libCurve_timeStampSplitCnt)
End Script

' 计算已分段 t 参数 (t∈[0, 1])
' @params timeStamp 当前时间戳
' @params curSplitIdx 当前段号
' @return 若当前 t 不属于当前段号则返回 -1, 否则返回范围为 [0, 1] 的参数
Export Script CUTimeCalcSplitTByStampAndIdx(timeStamp As Double, curSplitIdx As Long, Return Double)
    timeStamp = CUTimeCalcT(timeStamp)
    If curSplitIdx <= 0 Or curSplitIdx > libCurve_timeStampSplitCnt - 1 Then
        Return -1
    End If

    timeStamp = timeStamp * libCurve_timeStampSplitCnt
    If timeStamp < curSplitIdx Or timeStamp - curSplitIdx > 1.000000001 Then
        Return -1
    End If

    Return timeStamp - curSplitIdx
End Script

' ----------------------- 时间工具集 END
' ------------------------------------

' ------------------------------------
' ----------------------- Ease t BEGIN

Export Script CUEase_InSine(t As Double, Return Double)
    Return -Cos(t * 1.5707964) + 1
End Script

Export Script CUEase_OutSine(t As Double, Return Double)
    Return Sin(t * 1.5707964)
End Script

Export Script CUEase_InOutSine(t As Double, Return Double)
    Return -0.5 * (Cos(3.14159265357 * t) - 1)
End Script

Export Script CUEase_InQuad(t As Double, Return Double)
    Return t * t
End Script

Export Script CUEase_OutQuad(t As Double, Return Double)
    Return -t * (t - 2)
End Script

Export Script CUEase_InOutQuad(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return 0.5 * t * t
    End If
    t = t - 1
    Return -0.5 * (t * (t - 2) - 1)
End Script

Export Script CUEase_InCubic(t As Double, Return Double)
    Return t * t * t
End Script

Export Script CUEase_OutCubic(t As Double, Return Double)
    t = t - 1
    Return t * t * t + 1
End Script

Export Script CUEase_InOutCubic(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return 0.5 * t * t * t
    End If
    t = t - 2
    Return 0.5 * (t * t * t + 2)
End Script

Export Script CUEase_InQuart(t As Double, Return Double)
    Return t * t * t * t
End Script

Export Script CUEase_OutQuart(t As Double, Return Double)
    t = t - 1
    Return -t * t * t * t + 1
End Script

Export Script CUEase_InOutQuart(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return 0.5 * t * t * t * t
    End If
    t = t - 2
    Return -0.5 * (t * t * t * t - 2)
End Script

Export Script CUEase_InExpo(t As Double, Return Double)
    If Abs(t) < 0.000000001 Then
        Return 0
    End If
    Return 2 ^ (10 * (t - 1))
End Script

Export Script CUEase_OutExpo(t As Double, Return Double)
    If Abs(t - 1) < 0.000000001 Then
        Return 0
    End If
    Return - (2 ^ (-10 * t)) + 1
End Script

Export Script CUEase_InOutExpo(t As Double, Return Double)
    If Abs(t) < 0.000000001 Then
        Return 0
    End If
    If Abs(t - 1) < 0.000000001 Then
        Return 1
    End If

    t = t * 2
    If t < 1 Then
        Return 0.5 * (2 ^ (10 * t - 1))
    End If
    Return 0.5 * (2 ^ (-10 * (t - 1) + 2))
End Script

Export Script CUEase_InCirc(t As Double, Return Double)
    Return - Sqr(1 - t * t) - 1
End Script

Export Script CUEase_OutCirc(t As Double, Return Double)
    t = t - 1
    Return Sqr(1 - t * t)
End Script

Export Script CUEase_InOutCirc(t As Double, Return Double)
    t = t * 2
    If t < 1 Then
        Return - 0.5 * (Sqr(1 - t * t) - 1)
    End If
    t = t - 2
    Return 0.5 * (Sqr(1 - t * t) + 1)
End Script

' ----------------------- Ease t END
' ------------------------------------

' ------------------------------------
' ----------------------- Math END

Export Script CUMath_Max(a As Double, b As Double, Return Double)
    If a >= b Then
        Return a
    Else
        Return b
    End If
End Script

Export Script CUMath_Min(a As Double, b As Double, Return Double)
    If a <= b Then
        Return a
    Else
        Return b
    End If
End Script

Export Script CUMath_Clamp01(a As Double, Return Double)
    If a < 0 Then
        Return 0
    ElseIf a > 1 Then
        Return 1
    Else
        Return a
    End If
End Script

Export Script CUMath_Remap01(f As Double, t As Double, vv As Double, Return Double)
    If f = t Then
        Return vv
    End If
    Return (vv - f) / (t - f)
End Script

Export Script CUMath_Clamp(a As Double, b As Double, vv As Double, Return Double)
    If vv < a Then
        Return a
    ElseIf vv > b Then
        Return b
    Else
        Return vv
    End If
End Script

Export Script CUMath_Avg(a As Double, b As Double, Return Double)
    Return (a + b) / 2
End Script

Export Script CUMath_SinLerp(a As Double, b As Double, theta As Double, Return Double)
    theta = (Sin(theta * PI - 0.5 * PI) + 1) / 2
    Return a * (1 - theta) + theta * b
End Script

Export Script CUMath_Sin01(theta As Double, Return Double)
    Return (Sin(theta) + 1) / 2
End Script

Export Script CUMath_Cos01(theta As Double, Return Double)
    Return (Cos(theta) + 1) / 2
End Script

Export Script CUMath_Lerp(a As Double, b As Double, t As Double, Return Double)
    Return a * (1 - t) + t * b
End Script

Export Script CUMath_Step(a As Double, b As Double, Return Integer)
    If a < b Then
        Return 0
    Else
        Return 1
    End If
End Script

Export Script CUMath_Select(a As Double, b As Double, vv As Double, Return Double)
    If vv <= 0 Then
        Return a
    Else
        Return b
    End If
End Script

Export Script CUMath_SmoothStep(a As Double, b As Double, vv As Double, Return Double)
    If vv <= 0 Then
        Return a
    ElseIf vv >= 1 Then
        Return b
    Else
        Return CUMath_Lerp(a, b, vv * vv * (3 - 2 * vv))
    End If
End Script

Export Script CUMath_Frac(a As Double, Return Double)
    Return a - Int(a)
End Script

' 计算点乘
Export Script CUMath_Dot(a_x As Double, a_y As Double, b_x As Double, b_y As Double, Return Double)
    Return a_x * b_x + a_y * b_y
End Script

' 计算叉乘
Export Script CUMath_Cross(a_x As Double, a_y As Double, b_x As Double, b_y As Double, Return Double)
    Return a_x * b_y - a_y * b_x
End Script

' 计算反正切
' Return Atan(y / x)
Export Script CUMath_Atan2(y As Double, x As Double, Return Double)
    If x >= MIN_FLOAT Then
        Return Atn(y / x)
    ElseIf x <= -MIN_FLOAT And y >= 0 Then
        Return Atn(y / x) + PI
    ElseIf x <= -MIN_FLOAT And y < 0 Then
        Return Atn(y / x) - PI
    ElseIf MIN_FLOAT >= Abs(x) Then
        If y > 0 Then
            Return PI / 2
        ElseIf y < 0 Then
            Return -PI / 2
        End If
    End If
    Return 0
End Script

' 计算俩向量的逆时针夹角
Export Script CUMath_Angle(a_x As Double, a_y As Double, b_x As Double, b_y As Double, Return Double)
    Return CUMath_Atan2(CUMath_Cross(a_x, a_y, b_x, b_y), CUMath_Dot(a_x, a_y, b_x, b_y))
End Script

' 计算随机数
' @params seed 随机数种子
' @return 返回 [0, 1) 的随机数
Export Script CUMath_Hash(seed As Double, Return Double)
    seed = seed * 0.1031
    seed = seed - Int(seed)
    seed = seed * (seed + 33.33)
    seed = seed * (seed + seed)
    Return seed - Int(seed)
End Script

' 计算向量逆时针旋转
' angle 角度, 弧度制
Export Script CUMath_VecRotate(a_x As Double, a_y As Double, angle As Double, Return Integer)
    libCurve_retX = a_x * Cos(angle) + a_y * Sin(angle)
    libCurve_retY = -a_x * Sin(angle) + a_y * Cos(angle)
    Return libCurve_retX
End Script

' 计算向量归一化
' 返回值为 0 时表示向量长度为 0, 否则返回
Export Script CUMath_VecNormal(a_x As Double, a_y As Double, Return Integer)
    If a_x = 0 And a_y = 0 Then
        libCurve_retX = 0
        libCurve_retY = 0
        Return 0
    End If
    libCurve_retX = a_x / Sqr(a_x * a_x + a_y * a_y)
    libCurve_retY = a_y / Sqr(a_x * a_x + a_y * a_y)
    Return libCurve_retX
End Script

' 向量模长
Export Script CUMath_VecLength(a_x As Double, a_y As Double, Return Double)
    Return Sqr(a_x * a_x + a_y * a_y)
End Script

' 向量模长平方
Export Script CUMath_VecLength2(a_x As Double, a_y As Double, Return Double)
    Return a_x * a_x + a_y * a_y
End Script

' 向量加法
Export Script CUMath_VecAdd(a_x As Double, a_y As Double, b_x As Double, b_y As Double, Return Double)
    libCurve_retX = a_x + b_x
    libCurve_retY = a_y + b_y
    Return libCurve_retX
End Script

' 向量减法
Export Script CUMath_VecSub(a_x As Double, a_y As Double, b_x As Double, b_y As Double, Return Double)
    libCurve_retX = a_x - b_x
    libCurve_retY = a_y - b_y
    Return libCurve_retX
End Script

' 向量数乘
Export Script CUMath_VecMul(a_x As Double, a_y As Double, b As Double, Return Double)
    libCurve_retX = a_x * b
    libCurve_retY = a_y * b
    Return libCurve_retX
End Script

' 获取向量计算结果返回值 X
Export Script CUMath_GetVecRetX(Return Double)
    Return libCurve_retX
End Script

' 获取向量计算结果返回值 Y
Export Script CUMath_GetVecRetY(Return Double)
    Return libCurve_retY
End Script

' ----------------------- Math END
' ------------------------------------

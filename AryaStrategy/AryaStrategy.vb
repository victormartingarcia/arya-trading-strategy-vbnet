Imports System
Imports System.Collections.Generic
Imports TradingMotion.SDK.Algorithms
Imports TradingMotion.SDK.Algorithms.InputParameters
Imports TradingMotion.SDK.Markets.Charts
Imports TradingMotion.SDK.Markets.Orders
Imports TradingMotion.SDK.Markets.Indicators.Momentum
Imports TradingMotion.SDK.Markets.Indicators.OverlapStudies

Namespace AryaStrategy

    ''' <summary>
    ''' Arya trading rules:
    '''     * Entry: Stochastic %D indicator breaks above an upper bound (buy signal) or below a lower bound (sell signal)
    '''     * Exit: Trailing stop based on the entry price and moving according to price raise, or price reaches Profit target
    '''     * Filters: Day-of-week trading enabled, trading timeframe, volatility filter, ADX minimum level and bullish/bearish trend
    ''' </summary>
    Public Class AryaStrategy
        Inherits Strategy

        Dim stochasticIndicator As StochasticIndicator
        Dim adxIndicator As ADXIndicator
        Dim smaIndicator As SMAIndicator

        Dim trailingStopOrder As Order
        Dim profitOrder As Order

        Dim acceleration As Decimal
        Dim furthestClose As Decimal

        Public Sub New(ByVal mainChart As Chart, ByVal secondaryCharts As List(Of Chart))
            MyBase.New(mainChart, secondaryCharts)
        End Sub

        ''' <summary>
        ''' Strategy Name
        ''' </summary>
        ''' <returns>The complete Name of the strategy</returns>
        Public Overrides ReadOnly Property Name As String
            Get
                Return "Arya Strategy"
            End Get
        End Property

        ''' <summary>
        ''' Security filter that ensures the Position will be closed at the end of the trading session.
        ''' </summary>
        Public Overrides ReadOnly Property ForceCloseIntradayPosition As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Security filter that sets a maximum open position size of 1 contract (either side)
        ''' </summary>
        Public Overrides ReadOnly Property MaxOpenPosition As UInteger
            Get
                Return 1
            End Get
        End Property

        ''' <summary>
        ''' This strategy uses the Advanced Order Management mode
        ''' </summary>
        Public Overrides ReadOnly Property UsesAdvancedOrderManagement As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Strategy Parameter definition
        ''' </summary>
        Public Overrides Function SetInputParameters() As InputParameterList

            Dim parameters As New InputParameterList()

            ' Day of week trading enabled filters (0 = disabled)
            parameters.Add(New InputParameter("Monday Trading Enabled", 1))
            parameters.Add(New InputParameter("Tuesday Trading Enabled", 1))
            parameters.Add(New InputParameter("Wednesday Trading Enabled", 0))
            parameters.Add(New InputParameter("Thursday Trading Enabled", 0))
            parameters.Add(New InputParameter("Friday Trading Enabled", 1))

            ' Session time filter (entries will be only placed during this time frame)
            parameters.Add(New InputParameter("Trading Time Start", New TimeSpan(18, 0, 0)))
            parameters.Add(New InputParameter("Trading Time End", New TimeSpan(6, 0, 0)))

            ' The previous N bars period used for calculating Price Range
            parameters.Add(New InputParameter("Range Calculation Period", 10))

            ' Minimum Volatility allowed for placing entries
            parameters.Add(New InputParameter("Minimum Range Filter", 0.002D))

            ' The previous N bars period ADX indicator will use
            parameters.Add(New InputParameter("ADX Period", 14))
            ' The previous N bars period SMA indicator will use
            parameters.Add(New InputParameter("SMA Period", 78))

            ' Minimum ADX value for placing long entries
            parameters.Add(New InputParameter("Min ADX Long Entry", 12D))
            ' Minimum ADX value for placing short entries
            parameters.Add(New InputParameter("Min ADX Short Entry", 12D))

            ' The distance between the entry and the initial trailing stop price
            parameters.Add(New InputParameter("Trailing Stop Loss ticks distance", 24))
            ' The initial acceleration of the trailing stop
            parameters.Add(New InputParameter("Trailing Stop acceleration", 0.2D))

            ' The distance between the entry and the profit target level
            parameters.Add(New InputParameter("Profit Target ticks distance", 77))

            ' The previous N bars period Stochastic indicator will use
            parameters.Add(New InputParameter("Stochastic Period", 68))
            ' Break level of Stochastic %D we consider a buy signal
            parameters.Add(New InputParameter("Trend-following buy signal", 51D))
            ' Break level of Stochastic %D we consider a sell signal
            parameters.Add(New InputParameter("Trend-following sell signal", 49D))

            Return parameters

        End Function

        ''' <summary>
        ''' Initialization method
        ''' </summary>
        Public Overrides Sub OnInitialize()

            log.Debug("AryaStrategy onInitialize()")

            ' Adding a Stochastic indicator to strategy
            ' (see http://stockcharts.com/help/doku.php?id=chart_school:technical_indicators:stochastic_oscillato)
            stochasticIndicator = New StochasticIndicator(Bars.Bars, Me.GetInputParameter("Stochastic Period"))
            Me.AddIndicator("Stochastic Indicator", stochasticIndicator)

            ' Adding an ADX indicator to strategy
            ' (see http://www.investopedia.com/terms/a/adx.asp)
            adxIndicator = New ADXIndicator(Bars.Bars, Me.GetInputParameter("ADX Period"))
            Me.AddIndicator("ADX Indicator", adxIndicator)

            ' Adding a SMA indicator to strategy
            ' (see http://www.investopedia.com/terms/s/sma.asp)
            smaIndicator = New SMAIndicator(Bars.Close, Me.GetInputParameter("SMA Period"))
            Me.AddIndicator("SMA Indicator", smaIndicator)

        End Sub

        ''' <summary>
        ''' Strategy enter/exit/filtering rules
        ''' </summary>
        Public Overrides Sub OnNewBar()

            Dim buySignal As Decimal = Me.GetInputParameter("Trend-following buy signal")
            Dim sellSignal As Decimal = Me.GetInputParameter("Trend-following sell signal")

            Dim stopMargin As Decimal = Me.GetInputParameter("Trailing Stop Loss ticks distance") * Me.GetMainChart().Symbol.TickSize
            Dim profitMargin As Decimal = Me.GetInputParameter("Profit Target ticks distance") * Me.GetMainChart().Symbol.TickSize

            Dim longTradingEnabled As Boolean = False
            Dim shortTradingEnabled As Boolean = False

            ' Day-of-week filter
            If IsDayEnabledForTrading(Me.Bars.Time(0).DayOfWeek) Then

                ' Time-of-day filter
                If IsTimeEnabledForTrading(Me.Bars.Time(0)) Then

                    ' Volatility filter
                    If CalculateVolatilityRange() > Me.GetInputParameter("Minimum Range Filter") Then

                        ' ADX minimum level and current trending filters
                        If Me.GetOpenPosition() = 0 And IsADXEnabledForLongEntry() And IsBullishUnderlyingTrend() Then

                            longTradingEnabled = True

                        ElseIf Me.GetOpenPosition() = 0 And IsADXEnabledForShortEntry() And IsBearishUnderlyingTrend() Then

                            shortTradingEnabled = True

                        End If

                    End If

                End If

            End If

            If longTradingEnabled And stochasticIndicator.GetD()(1) <= buySignal And stochasticIndicator.GetD()(0) > buySignal Then

                ' BUY SIGNAL: Stochastic %D crosses above "buy signal" level
                Dim buyOrder As MarketOrder = New MarketOrder(OrderSide.Buy, 1, "Enter long position")
                Me.InsertOrder(buyOrder)

                trailingStopOrder = New StopOrder(OrderSide.Sell, 1, Me.Bars.Close(0) - stopMargin, "Catastrophic stop long exit")
                Me.InsertOrder(trailingStopOrder)

                profitOrder = New LimitOrder(OrderSide.Sell, 1, Me.Bars.Close(0) + profitMargin, "Profit stop long exit")
                Me.InsertOrder(profitOrder)

                ' Linking Stop and Limit orders: when one is executed, the other is cancelled
                trailingStopOrder.IsChildOf = profitOrder
                profitOrder.IsChildOf = trailingStopOrder

                ' Setting the initial acceleration for the trailing stop and the furthest (the most extreme) close price
                acceleration = Me.GetInputParameter("Trailing Stop acceleration")
                furthestClose = Me.Bars.Close(0)

            ElseIf shortTradingEnabled And stochasticIndicator.GetD()(1) >= sellSignal And stochasticIndicator.GetD()(0) < sellSignal Then

                ' SELL SIGNAL: Stochastic %D crosses below "sell signal" level
                Dim sellOrder As MarketOrder = New MarketOrder(OrderSide.Sell, 1, "Enter short position")
                Me.InsertOrder(sellOrder)

                trailingStopOrder = New StopOrder(OrderSide.Buy, 1, Me.Bars.Close(0) + stopMargin, "Catastrophic stop short exit")
                Me.InsertOrder(trailingStopOrder)

                profitOrder = New LimitOrder(OrderSide.Buy, 1, Me.Bars.Close(0) - profitMargin, "Profit stop short exit")
                Me.InsertOrder(profitOrder)

                ' Linking Stop and Limit orders: when one is executed, the other is cancelled
                trailingStopOrder.IsChildOf = profitOrder
                profitOrder.IsChildOf = trailingStopOrder

                ' Setting the initial acceleration for the trailing stop and the furthest (the most extreme) close price
                acceleration = Me.GetInputParameter("Trailing Stop acceleration")
                furthestClose = Me.Bars.Close(0)

            ElseIf Me.GetOpenPosition() = 1 And Me.Bars.Close(0) > furthestClose Then

                ' We're long and the price has moved in our favour

                furthestClose = Me.Bars.Close(0)

                ' Increasing acceleration
                acceleration = acceleration * (furthestClose - trailingStopOrder.Price)

                ' Checking if trailing the stop order would exceed the current market price
                If trailingStopOrder.Price + acceleration < Me.Bars.Close(0) Then

                    ' Setting the new price for the trailing stop
                    trailingStopOrder.Price = trailingStopOrder.Price + acceleration
                    trailingStopOrder.Label = "Trailing stop long exit"
                    Me.ModifyOrder(trailingStopOrder)

                Else

                    ' Cancelling the order and closing the position
                    Me.CancelOrder(trailingStopOrder)
                    Me.CancelOrder(profitOrder)

                    Dim exitLongOrder As MarketOrder = New MarketOrder(OrderSide.Sell, 1, "Exit long position")
                    Me.InsertOrder(exitLongOrder)

                End If

            ElseIf Me.GetOpenPosition() = -1 And Me.Bars.Close(0) < furthestClose Then

                ' We're short and the price has moved in our favour

                furthestClose = Me.Bars.Close(0)

                ' Increasing acceleration
                acceleration = acceleration * Math.Abs(trailingStopOrder.Price - furthestClose)

                ' Checking if trailing the stop order would exceed the current market price
                If trailingStopOrder.Price - acceleration > Me.Bars.Close(0) Then

                    ' Setting the new price for the trailing stop
                    trailingStopOrder.Price = trailingStopOrder.Price - acceleration
                    trailingStopOrder.Label = "Trailing stop short exit"
                    Me.ModifyOrder(trailingStopOrder)

                Else

                    ' Cancelling the order and closing the position
                    Me.CancelOrder(trailingStopOrder)
                    Me.CancelOrder(profitOrder)

                    Dim exitShortOrder As MarketOrder = New MarketOrder(OrderSide.Buy, 1, "Exit short position")
                    Me.InsertOrder(exitShortOrder)

                End If

            End If

        End Sub

        Private Function IsDayEnabledForTrading(currentDay As DayOfWeek) As Boolean

            ' Check if current day is available to trade
            If currentDay = DayOfWeek.Monday And Me.GetInputParameter("Monday Trading Enabled") = 0 Then
                Return False
            End If

            If currentDay = DayOfWeek.Tuesday And Me.GetInputParameter("Tuesday Trading Enabled") = 0 Then
                Return False
            End If

            If currentDay = DayOfWeek.Wednesday And Me.GetInputParameter("Wednesday Trading Enabled") = 0 Then
                Return False
            End If

            If currentDay = DayOfWeek.Thursday And Me.GetInputParameter("Thursday Trading Enabled") = 0 Then
                Return False
            End If

            If currentDay = DayOfWeek.Friday And Me.GetInputParameter("Friday Trading Enabled") = 0 Then
                Return False
            End If

            Return True

        End Function

        Private Function IsTimeEnabledForTrading(currentBar As DateTime) As Boolean

            ' Check if the current bar's time of day is inside the enabled time range to trade
            Dim tradingTimeStart As TimeSpan = Me.GetInputParameter("Trading Time Start")
            Dim tradingTimeEnd As TimeSpan = Me.GetInputParameter("Trading Time End")

            If tradingTimeStart <= tradingTimeEnd Then

                Return currentBar.TimeOfDay >= tradingTimeStart And
                    currentBar.TimeOfDay <= tradingTimeEnd

            Else

                Return currentBar.TimeOfDay >= tradingTimeStart Or
                    currentBar.TimeOfDay <= tradingTimeEnd

            End If

        End Function

        Private Function CalculateVolatilityRange() As Decimal

            ' Set current Range as Max(High) - Min(Low) of last "Range Calculation Period" bars
            Dim lookbackPeriod As Integer = Me.GetInputParameter("Range Calculation Period")
            Return Me.Bars.High.GetHighestValue(lookbackPeriod) - Me.Bars.Low.GetLowestValue(lookbackPeriod)

        End Function

        Private Function IsADXEnabledForLongEntry() As Boolean

            ' Long entry enabled if ADX is above "Min ADX Long Entry" parameter
            Return adxIndicator.GetADX()(0) >= Me.GetInputParameter("Min ADX Long Entry")

        End Function

        Private Function IsBullishUnderlyingTrend() As Boolean

            ' Consider bullish underlying trend if SMA has raised on the last bar
            Return smaIndicator.GetAvSimple()(0) > smaIndicator.GetAvSimple()(1)

        End Function

        Private Function IsADXEnabledForShortEntry() As Boolean

            ' Short entry enabled if ADX is above "Min ADX Short Entry" parameter
            Return adxIndicator.GetADX()(0) >= Me.GetInputParameter("Min ADX Short Entry")

        End Function

        Private Function IsBearishUnderlyingTrend() As Boolean

            ' Consider bearish underlying trend if SMA has fallen on the last bar
            Return smaIndicator.GetAvSimple()(0) < smaIndicator.GetAvSimple()(1)

        End Function

    End Class
End Namespace
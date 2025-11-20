"""
Black-Scholes Option Pricing Model
Python implementation based on black_scholes.luau
"""

import math
from typing import Literal
import xllify


def norm_cdf(x: float) -> float:
    """
    Normal CDF approximation using Abramowitz and Stegun method.
    """
    if x < -7:
        return 0.0
    elif x > 7:
        return 1.0

    sign = 1
    if x < 0:
        sign = -1
        x = -x

    t = 1 / (1 + 0.2316419 * x)
    y = t * (
        0.31938153
        + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429)))
    )
    z = math.exp(-x * x / 2) / math.sqrt(2 * math.pi)

    if sign == 1:
        return 1 - z * y
    else:
        return z * y


def black_scholes(
    S: float,
    K: float,
    T: float,
    r: float,
    sigma: float,
    option_type: Literal["call", "put"],
) -> float:
    """
    Black-Scholes formula for call and put options.

    Args:
        S: Current stock price
        K: Strike price
        T: Time to expiration (years)
        r: Risk-free interest rate
        sigma: Volatility (annualized)
        option_type: "call" or "put"

    Returns:
        Option price
    """
    # Handle expired options (return intrinsic value)
    if T <= 0:
        if option_type == "call":
            return max(S - K, 0)
        else:
            return max(K - S, 0)

    # Calculate d1 and d2
    d1 = (math.log(S / K) + (r + sigma * sigma / 2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)

    # Calculate option price
    if option_type == "call":
        return S * norm_cdf(d1) - K * math.exp(-r * T) * norm_cdf(d2)
    else:  # put
        return K * math.exp(-r * T) * norm_cdf(-d2) - S * norm_cdf(-d1)


@xllify.fn("xllipy.BSCall")
def call_price(S: float, K: float, T: float, r: float, sigma: float) -> float:
    """
    Calculate Black-Scholes call option price.

    Args:
        S: Current stock price
        K: Strike price
        T: Time to expiration (years)
        r: Risk-free interest rate
        sigma: Volatility (annualized)

    Returns:
        Call option price

    Example:
        >>> call_price(100, 100, 1, 0.05, 0.20)
        10.450583572185565
    """
    return black_scholes(S, K, T, r, sigma, "call")


@xllify.fn("xllipy.BSPut")
def put_price(S: float, K: float, T: float, r: float, sigma: float) -> float:
    """
    Calculate Black-Scholes put option price.
    """
    return black_scholes(S, K, T, r, sigma, "put")


def market_state(
    call_price: float, put_price: float, S: float, K: float, T: float, r: float
) -> str:
    """
    Determine market state from Black-Scholes call and put prices.
    Analyzes moneyness and checks for put-call parity violations.
    """
    # Calculate put-call parity: C - P = S - K*e^(-rT)
    theoretical_diff = S - K * math.exp(-r * T)
    actual_diff = call_price - put_price
    parity_deviation = abs(actual_diff - theoretical_diff)

    # Determine moneyness
    moneyness = S / K
    state = ""

    if moneyness > 1.05:
        state = "Deep ITM Call / OTM Put"
    elif moneyness > 1.01:
        state = "ITM Call / OTM Put"
    elif moneyness >= 0.99:
        state = "ATM"
    elif moneyness >= 0.95:
        state = "OTM Call / ITM Put"
    else:
        state = "OTM Call / Deep ITM Put"

    # Check put-call parity violation (more than 1% deviation)
    if parity_deviation / S > 0.01:
        state = state + " [Parity Violation]"

    return state

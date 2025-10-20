-- Test Suite for black_scholes.luau
-- Tests the Black-Scholes option pricing functions: bs.Call and bs.Put

xllify.TestSuite("Black-Scholes Option Pricing", function()

    -- =============================================================================
    -- bs.Call Tests
    -- =============================================================================

    xllify.Test("Call should price at-the-money option correctly", function()
        -- S=K=100, T=1 year, r=5%, sigma=20%
        local result = xllify.Call("bs.Call", 100, 100, 1, 0.05, 0.2)
        -- Expected value approximately 10.45
        xllify.Assert.IsNumber(result)
        xllify.Assert.True(result > 10 and result < 11, "ATM call should be ~10.45")
    end)

    xllify.Test("Call should price in-the-money option higher than ATM", function()
        -- ITM: S > K
        local itm = xllify.Call("bs.Call", 110, 100, 1, 0.05, 0.2)
        local atm = xllify.Call("bs.Call", 100, 100, 1, 0.05, 0.2)
        xllify.Assert.True(itm > atm, "ITM call should be worth more than ATM")
    end)

    xllify.Test("Call should price out-of-the-money option lower than ATM", function()
        -- OTM: S < K
        local otm = xllify.Call("bs.Call", 90, 100, 1, 0.05, 0.2)
        local atm = xllify.Call("bs.Call", 100, 100, 1, 0.05, 0.2)
        xllify.Assert.True(otm < atm, "OTM call should be worth less than ATM")
    end)

    xllify.Test("Call should have intrinsic value at expiration (T=0, ITM)", function()
        -- At expiration, call = max(S - K, 0)
        local result = xllify.Call("bs.Call", 110, 100, 0, 0.05, 0.2)
        xllify.Assert.Equal(10, result)
    end)

    xllify.Test("Call should be worthless at expiration (T=0, OTM)", function()
        -- At expiration, OTM call = 0
        local result = xllify.Call("bs.Call", 90, 100, 0, 0.05, 0.2)
        xllify.Assert.Equal(0, result)
    end)

    xllify.Test("Call should increase with higher volatility", function()
        local lowVol = xllify.Call("bs.Call", 100, 100, 1, 0.05, 0.1)
        local highVol = xllify.Call("bs.Call", 100, 100, 1, 0.05, 0.3)
        xllify.Assert.True(highVol > lowVol, "Higher volatility should increase call value")
    end)

    xllify.Test("Call should increase with longer time to expiration", function()
        local shortTime = xllify.Call("bs.Call", 100, 100, 0.5, 0.05, 0.2)
        local longTime = xllify.Call("bs.Call", 100, 100, 2, 0.05, 0.2)
        xllify.Assert.True(longTime > shortTime, "Longer time should increase call value")
    end)

    xllify.Test("Call should increase with higher risk-free rate", function()
        local lowRate = xllify.Call("bs.Call", 100, 100, 1, 0.01, 0.2)
        local highRate = xllify.Call("bs.Call", 100, 100, 1, 0.10, 0.2)
        xllify.Assert.True(highRate > lowRate, "Higher rate should increase call value")
    end)

    xllify.Test("Call should always be non-negative", function()
        local result = xllify.Call("bs.Call", 50, 100, 1, 0.05, 0.2)
        xllify.Assert.True(result >= 0, "Call option cannot have negative value")
    end)

    xllify.Test("Call should handle deep in-the-money option", function()
        -- Very deep ITM: S = 200, K = 100
        local result = xllify.Call("bs.Call", 200, 100, 1, 0.05, 0.2)
        -- Should be close to S - K*exp(-rT) = 200 - 100*exp(-0.05) ≈ 104.88
        xllify.Assert.True(result > 100, "Deep ITM call should be > 100")
    end)

    -- =============================================================================
    -- bs.Put Tests
    -- =============================================================================

    xllify.Test("Put should price at-the-money option correctly", function()
        -- S=K=100, T=1 year, r=5%, sigma=20%
        local result = xllify.Call("bs.Put", 100, 100, 1, 0.05, 0.2)
        -- Expected value approximately 5.57
        xllify.Assert.IsNumber(result)
        xllify.Assert.True(result > 5 and result < 6, "ATM put should be ~5.57")
    end)

    xllify.Test("Put should price in-the-money option higher than ATM", function()
        -- ITM for put: S < K
        local itm = xllify.Call("bs.Put", 90, 100, 1, 0.05, 0.2)
        local atm = xllify.Call("bs.Put", 100, 100, 1, 0.05, 0.2)
        xllify.Assert.True(itm > atm, "ITM put should be worth more than ATM")
    end)

    xllify.Test("Put should price out-of-the-money option lower than ATM", function()
        -- OTM for put: S > K
        local otm = xllify.Call("bs.Put", 110, 100, 1, 0.05, 0.2)
        local atm = xllify.Call("bs.Put", 100, 100, 1, 0.05, 0.2)
        xllify.Assert.True(otm < atm, "OTM put should be worth less than ATM")
    end)

    xllify.Test("Put should have intrinsic value at expiration (T=0, ITM)", function()
        -- At expiration, put = max(K - S, 0)
        local result = xllify.Call("bs.Put", 90, 100, 0, 0.05, 0.2)
        xllify.Assert.Equal(10, result)
    end)

    xllify.Test("Put should be worthless at expiration (T=0, OTM)", function()
        -- At expiration, OTM put = 0
        local result = xllify.Call("bs.Put", 110, 100, 0, 0.05, 0.2)
        xllify.Assert.Equal(0, result)
    end)

    xllify.Test("Put should increase with higher volatility", function()
        local lowVol = xllify.Call("bs.Put", 100, 100, 1, 0.05, 0.1)
        local highVol = xllify.Call("bs.Put", 100, 100, 1, 0.05, 0.3)
        xllify.Assert.True(highVol > lowVol, "Higher volatility should increase put value")
    end)

    xllify.Test("Put should increase with longer time to expiration", function()
        local shortTime = xllify.Call("bs.Put", 100, 100, 0.5, 0.05, 0.2)
        local longTime = xllify.Call("bs.Put", 100, 100, 2, 0.05, 0.2)
        xllify.Assert.True(longTime > shortTime, "Longer time should increase put value")
    end)

    xllify.Test("Put should decrease with higher risk-free rate", function()
        local lowRate = xllify.Call("bs.Put", 100, 100, 1, 0.01, 0.2)
        local highRate = xllify.Call("bs.Put", 100, 100, 1, 0.10, 0.2)
        xllify.Assert.True(lowRate > highRate, "Higher rate should decrease put value")
    end)

    xllify.Test("Put should always be non-negative", function()
        local result = xllify.Call("bs.Put", 150, 100, 1, 0.05, 0.2)
        xllify.Assert.True(result >= 0, "Put option cannot have negative value")
    end)

    xllify.Test("Put should handle deep in-the-money option", function()
        -- Very deep ITM for put: S = 50, K = 100
        local result = xllify.Call("bs.Put", 50, 100, 1, 0.05, 0.2)
        -- Should be significant, close to K*exp(-rT) - S ≈ 100*exp(-0.05) - 50 ≈ 45.12
        xllify.Assert.True(result > 40, "Deep ITM put should be > 40")
    end)

    -- =============================================================================
    -- Put-Call Parity Tests
    -- =============================================================================

    xllify.Test("Put-Call parity should hold", function()
        -- Put-Call Parity: C - P = S - K*exp(-rT)
        local S, K, T, r, sigma = 100, 100, 1, 0.05, 0.2
        local call = xllify.Call("bs.Call", S, K, T, r, sigma)
        local put = xllify.Call("bs.Put", S, K, T, r, sigma)
        local parity = call - put
        local expected = S - K * math.exp(-r * T)
        -- Allow small numerical error (0.01)
        xllify.Assert.True(math.abs(parity - expected) < 0.01, "Put-Call parity should hold")
    end)

    xllify.Test("Put-Call parity should hold for ITM options", function()
        local S, K, T, r, sigma = 120, 100, 1, 0.05, 0.2
        local call = xllify.Call("bs.Call", S, K, T, r, sigma)
        local put = xllify.Call("bs.Put", S, K, T, r, sigma)
        local parity = call - put
        local expected = S - K * math.exp(-r * T)
        xllify.Assert.True(math.abs(parity - expected) < 0.01, "Put-Call parity should hold for ITM")
    end)

    xllify.Test("Put-Call parity should hold for OTM options", function()
        local S, K, T, r, sigma = 80, 100, 1, 0.05, 0.2
        local call = xllify.Call("bs.Call", S, K, T, r, sigma)
        local put = xllify.Call("bs.Put", S, K, T, r, sigma)
        local parity = call - put
        local expected = S - K * math.exp(-r * T)
        xllify.Assert.True(math.abs(parity - expected) < 0.01, "Put-Call parity should hold for OTM")
    end)

    -- =============================================================================
    -- Edge Cases and Boundary Conditions
    -- =============================================================================

    xllify.Test("Call should handle very small time to expiration", function()
        local result = xllify.Call("bs.Call", 105, 100, 0.001, 0.05, 0.2)
        xllify.Assert.True(result >= 5 and result < 6, "Very short-dated ITM call ~ intrinsic")
    end)

    xllify.Test("Put should handle very small time to expiration", function()
        local result = xllify.Call("bs.Put", 95, 100, 0.001, 0.05, 0.2)
        xllify.Assert.True(result >= 4.9 and result < 5.1, "Very short-dated ITM put ~ intrinsic")
    end)

    xllify.Test("Call should handle very low volatility", function()
        local result = xllify.Call("bs.Call", 100, 100, 1, 0.05, 0.01)
        xllify.Assert.IsNumber(result)
        xllify.Assert.True(result >= 0, "Low volatility call should still be valid")
    end)

    xllify.Test("Put should handle very low volatility", function()
        local result = xllify.Call("bs.Put", 100, 100, 1, 0.05, 0.01)
        xllify.Assert.IsNumber(result)
        xllify.Assert.True(result >= 0, "Low volatility put should still be valid")
    end)

    xllify.Test("Call should handle zero risk-free rate", function()
        local result = xllify.Call("bs.Call", 100, 100, 1, 0, 0.2)
        xllify.Assert.IsNumber(result)
        xllify.Assert.True(result > 0, "Call with r=0 should have positive value")
    end)

    xllify.Test("Put should handle zero risk-free rate", function()
        local result = xllify.Call("bs.Put", 100, 100, 1, 0, 0.2)
        xllify.Assert.IsNumber(result)
        xllify.Assert.True(result > 0, "Put with r=0 should have positive value")
    end)

end)
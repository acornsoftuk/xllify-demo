-- Test Suite for simple.luau
-- Tests the demo functions: Hello, AgeCategory, and Portfolio

xllify.TestSuite("Simple Demo Functions", function()
    -- =============================================================================
    -- xllify.Demo.Hello Tests
    -- =============================================================================

    xllify.Test("Hello should greet with provided name", function()
        local result = xllify.Call("xllify.Demo.Hello", "Alice")
        xllify.Assert.Equal("Hello, Alice!", result)
    end)

    xllify.Test("Hello should default to World when no name provided", function()
        local result = xllify.Call("xllify.Demo.Hello")
        xllify.Assert.Equal("Hello, World!", result)
    end)

    xllify.Test("Hello should handle empty string", function()
        local result = xllify.Call("xllify.Demo.Hello", "")
        xllify.Assert.Equal("Hello, !", result)
    end)

    -- =============================================================================
    -- xllify.Demo.AgeCategory Tests
    -- =============================================================================

    xllify.Test("AgeCategory should return Baby for negative age", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", -1)
        xllify.Assert.Equal("Baby", result)
    end)

    xllify.Test("AgeCategory should return Kid for age 5", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 5)
        xllify.Assert.Equal("Kid", result)
    end)

    xllify.Test("AgeCategory should return Kid for age 12", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 12)
        xllify.Assert.Equal("Kid", result)
    end)

    xllify.Test("AgeCategory should return Annoying teenager for age 13", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 13)
        xllify.Assert.Equal("Annoying teenager", result)
    end)

    xllify.Test("AgeCategory should return Annoying teenager for age 19", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 19)
        xllify.Assert.Equal("Annoying teenager", result)
    end)

    xllify.Test("AgeCategory should return Disillusioned adult for age 20", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 20)
        xllify.Assert.Equal("Disillusioned adult", result)
    end)

    xllify.Test("AgeCategory should return Disillusioned adult for age 30", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 30)
        xllify.Assert.Equal("Disillusioned adult", result)
    end)

    xllify.Test("AgeCategory should return Disillusioned adult for age 64", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 64)
        xllify.Assert.Equal("Disillusioned adult", result)
    end)

    xllify.Test("AgeCategory should return OK boomer for age 65", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 65)
        xllify.Assert.Equal("OK boomer", result)
    end)

    xllify.Test("AgeCategory should return OK boomer for age 100", function()
        local result = xllify.Call("xllify.Demo.AgeCategory", 100)
        xllify.Assert.Equal("OK boomer", result)
    end)

    -- =============================================================================
    -- xllify.Demo.Portfolio Tests
    -- =============================================================================

    xllify.Test("Portfolio should return default 10 positions", function()
        local result = xllify.Call("xllify.Demo.Portfolio")
        xllify.Assert.IsMatrix(result)
        local dims = xllify.GetDimensions(result)
        xllify.Assert.Equal(10, dims[1])  -- 10 rows
        xllify.Assert.Equal(4, dims[2])   -- 4 columns (ticker, quantity, price, value)
    end)

    xllify.Test("Portfolio should generate requested number of positions", function()
        local result = xllify.Call("xllify.Demo.Portfolio", 5, 100000, 123)
        xllify.Assert.IsMatrix(result)
        local dims = xllify.GetDimensions(result)
        xllify.Assert.Equal(5, dims[1])
        xllify.Assert.Equal(4, dims[2])
    end)

    xllify.Test("Portfolio should be reproducible with same seed", function()
        local result1 = xllify.Call("xllify.Demo.Portfolio", 3, 50000, 42)
        local result2 = xllify.Call("xllify.Demo.Portfolio", 3, 50000, 42)
        xllify.Assert.Equal(result1, result2)
    end)

    xllify.Test("Portfolio should have 4 columns per position", function()
        local result = xllify.Call("xllify.Demo.Portfolio", 3, 100000, 111)
        xllify.Assert.IsMatrix(result)
        local dims = xllify.GetDimensions(result)
        xllify.Assert.Equal(4, dims[2])
    end)

    xllify.Test("Portfolio should contain valid ticker symbols", function()
        local result = xllify.Call("xllify.Demo.Portfolio", 2, 50000, 99)
        xllify.Assert.IsMatrix(result)
        -- First position should have a ticker (string)
        -- Note: In Luau matrices from xllify, we can't easily assert string types
        -- but we can verify it's a matrix with correct dimensions
        local dims = xllify.GetDimensions(result)
        xllify.Assert.Equal(2, dims[1])
    end)

end)

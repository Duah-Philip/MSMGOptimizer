test_that("ShinyMSMGOptimizer function exists", {
  expect_true(is.function(ShinyMSMGOptimizer))
})

test_that("shiny app directory exists", {
  app_dir <- system.file("shiny", package = "MSMGOptimizer")
  expect_true(nchar(app_dir) > 0)
})

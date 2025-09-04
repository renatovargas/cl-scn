aggregate_column_names <- function(df, column_map) {
  # Aggregate columns
  df <- aggregate(t(df), list(agg1 = column_map$new_name), FUN = sum)
  rownames(df) <- df$agg1
  df <- t(df[, -1]) # Quitar la columna de agregación y transponer de vuelta
  return(as.data.frame(df))
}

aggregate_row_names <- function(df, row_map) {
  # Transpose for row aggregation
  df <- aggregate(df, list(agg1 = row_map$new_name), FUN = sum)
  rownames(df) <- df$agg1
  df <- df[, -1] # Quitar la columna de agreagación
  return(as.data.frame(df))
}

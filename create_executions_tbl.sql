CREATE TABLE "Executions" (
  trade_id INT,
  account_number INT,
  order_number INT,
  ticker VARCHAR(6),
  action VARCHAR(4),
  time_of_execution TIME,
  quantity_executed INT,
  price_executed NUMERIC,
  PRIMARY KEY(account_number, order_number, time_of_execution)
);
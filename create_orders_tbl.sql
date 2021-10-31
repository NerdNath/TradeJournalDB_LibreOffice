CREATE TABLE "Orders" (
  trade_id INT,
  account_number INT,
  order_number INT,
  ticker VARCHAR(6),
  action VARCHAR(4),
  submit_date DATE,
  exe_type VARCHAR(3),
  requested_price NUMERIC,
  status VARCHAR(20),
  qty_requested INT,
  PRIMARY KEY(account_number, order_number),
  CONSTRAINT UC_order UNIQUE(account_number,order_number,ticker,action)
);
// ============================================================
// IQBAL TRADERS — مکمل یک فائل ورژن (Merged main.dart)
// تمام Models، Database، Provider، Widgets، Screens ایک فائل میں
// ============================================================

import 'dart:convert';
import 'dart:io';

import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:provider/provider.dart';
import 'package:intl/intl.dart';
import 'package:sqflite/sqflite.dart';
import 'package:path/path.dart' as dbPath;
import 'package:path_provider/path_provider.dart';
import 'package:file_picker/file_picker.dart';
import 'package:share_plus/share_plus.dart';
import 'package:excel/excel.dart' hide Border;
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:printing/printing.dart';

// ============================================================
// MAIN ENTRY POINT
// ============================================================

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  configureGoogleFonts();

  await SystemChrome.setPreferredOrientations([
    DeviceOrientation.portraitUp,
    DeviceOrientation.portraitDown,
  ]);

  SystemChrome.setSystemUIOverlayStyle(const SystemUiOverlayStyle(
    statusBarColor: Color(0xFF5a67d8),
    statusBarIconBrightness: Brightness.light,
  ));

  runApp(
    ChangeNotifierProvider(
      create: (_) => BusinessProvider()..initialize(),
      child: const IqbalTradersApp(),
    ),
  );
}

// ============================================================
// CONSTANTS
// ============================================================

class AppConstants {
  static const String appName    = 'IQBAL TRADERS';
  static const String appVersion = '1.0.0';
  static const String dbName     = 'iqbal_traders.db';
  static const int    dbVersion  = 1;

  static const List<String> expenseCategories = [
    'کرایہ','بجلی','پانی','گیس','تنخواہ','ٹرانسپورٹ',
    'مرمت','اشتہار','ٹیکس','بیمہ','دفتری اخراجات',
    'صفائی','سیکیورٹی','ٹیلیفون','انٹرنیٹ',
    'پیکنگ','لوڈنگ/انلوڈنگ','دیگر',
  ];

  static const List<String> stockUnits = [
    'KG','Gram','Ton','Bag','Box','Carton',
    'Piece','Liter','Meter','Yard','Dozen',
  ];
}

class DbTables {
  static const settings             = 'settings';
  static const customers            = 'customers';
  static const customerTransactions = 'customer_transactions';
  static const suppliers            = 'suppliers';
  static const supplierTransactions = 'supplier_transactions';
  static const udhar                = 'udhar';
  static const udharTransactions    = 'udhar_transactions';
  static const stock                = 'stock';
  static const purchaseHistory      = 'purchase_history';
  static const stockAdjustments     = 'stock_adjustments';
  static const sales                = 'sales';
  static const saleItems            = 'sale_items';
  static const purchases            = 'purchases';
  static const purchaseItems        = 'purchase_items';
  static const expenses             = 'expenses';
  static const cashInOut            = 'cash_in_out';
  static const cashLedger           = 'cash_ledger';
  static const partners             = 'partners';
  static const partnerWithdrawals   = 'partner_withdrawals';
}

// ============================================================
// IQBAL TRADERS - All Data Models (Null-Safe, Clean)
// ============================================================

class AppTransaction {
  final String date;
  final String type;
  final double amount;
  final String note;
  final int timestamp;
  final int? saleId;
  final int? purchaseId;

  const AppTransaction({
    required this.date,
    required this.type,
    required this.amount,
    this.note = '',
    required this.timestamp,
    this.saleId,
    this.purchaseId,
  });

  Map<String, dynamic> toMap() => {
    'date': date, 'type': type, 'amount': amount,
    'note': note, 'timestamp': timestamp,
    'saleId': saleId, 'purchaseId': purchaseId,
  };

  factory AppTransaction.fromMap(Map<String, dynamic> m) => AppTransaction(
    date: (m['date'] as String?) ?? '',
    type: (m['type'] as String?) ?? '',
    amount: (m['amount'] as num?)?.toDouble() ?? 0.0,
    note: (m['note'] as String?) ?? '',
    timestamp: (m['timestamp'] as int?) ?? 0,
    saleId: m['saleId'] as int?,
    purchaseId: m['purchaseId'] as int?,
  );
}

class Customer {
  final int id;
  String name;
  String phone;
  String address;
  double balance;
  List<AppTransaction> transactions;
  final String createdAt;

  Customer({
    required this.id, required this.name, required this.phone,
    required this.address, required this.balance,
    required this.transactions, required this.createdAt,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'name': name, 'phone': phone,
    'address': address, 'balance': balance, 'createdAt': createdAt,
  };

  factory Customer.fromMap(Map<String, dynamic> m) => Customer(
    id: (m['id'] as int?) ?? 0,
    name: (m['name'] as String?) ?? '',
    phone: (m['phone'] as String?) ?? '',
    address: (m['address'] as String?) ?? '',
    balance: (m['balance'] as num?)?.toDouble() ?? 0.0,
    transactions: [],
    createdAt: (m['createdAt'] as String?) ?? DateTime.now().toIso8601String(),
  );

  Customer copyWith({String? name, String? phone, String? address,
      double? balance, List<AppTransaction>? transactions}) =>
    Customer(id: id, name: name ?? this.name, phone: phone ?? this.phone,
      address: address ?? this.address, balance: balance ?? this.balance,
      transactions: transactions ?? this.transactions, createdAt: createdAt);
}

class Supplier {
  final int id;
  String name;
  String phone;
  String address;
  double balance;
  List<AppTransaction> transactions;
  final String createdAt;

  Supplier({
    required this.id, required this.name, required this.phone,
    required this.address, required this.balance,
    required this.transactions, required this.createdAt,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'name': name, 'phone': phone,
    'address': address, 'balance': balance, 'createdAt': createdAt,
  };

  factory Supplier.fromMap(Map<String, dynamic> m) => Supplier(
    id: (m['id'] as int?) ?? 0,
    name: (m['name'] as String?) ?? '',
    phone: (m['phone'] as String?) ?? '',
    address: (m['address'] as String?) ?? '',
    balance: (m['balance'] as num?)?.toDouble() ?? 0.0,
    transactions: [],
    createdAt: (m['createdAt'] as String?) ?? DateTime.now().toIso8601String(),
  );

  Supplier copyWith({String? name, String? phone, String? address,
      double? balance, List<AppTransaction>? transactions}) =>
    Supplier(id: id, name: name ?? this.name, phone: phone ?? this.phone,
      address: address ?? this.address, balance: balance ?? this.balance,
      transactions: transactions ?? this.transactions, createdAt: createdAt);
}

class UdharPerson {
  final int id;
  String name;
  String phone;
  String address;
  double balance;
  List<AppTransaction> transactions;
  final String createdAt;

  UdharPerson({
    required this.id, required this.name, required this.phone,
    required this.address, required this.balance,
    required this.transactions, required this.createdAt,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'name': name, 'phone': phone,
    'address': address, 'balance': balance, 'createdAt': createdAt,
  };

  factory UdharPerson.fromMap(Map<String, dynamic> m) => UdharPerson(
    id: (m['id'] as int?) ?? 0,
    name: (m['name'] as String?) ?? '',
    phone: (m['phone'] as String?) ?? '',
    address: (m['address'] as String?) ?? '',
    balance: (m['balance'] as num?)?.toDouble() ?? 0.0,
    transactions: [],
    createdAt: (m['createdAt'] as String?) ?? DateTime.now().toIso8601String(),
  );

  UdharPerson copyWith({String? name, String? phone, String? address,
      double? balance, List<AppTransaction>? transactions}) =>
    UdharPerson(id: id, name: name ?? this.name, phone: phone ?? this.phone,
      address: address ?? this.address, balance: balance ?? this.balance,
      transactions: transactions ?? this.transactions, createdAt: createdAt);
}

class PurchaseHistory {
  final int supplierId;
  final double rate;
  final double quantity;
  final String date;
  final int? purchaseId;
  final int timestamp;

  const PurchaseHistory({
    required this.supplierId, required this.rate,
    required this.quantity, required this.date,
    this.purchaseId, required this.timestamp,
  });

  Map<String, dynamic> toMap() => {
    'supplierId': supplierId, 'rate': rate, 'quantity': quantity,
    'date': date, 'purchaseId': purchaseId, 'timestamp': timestamp,
  };

  factory PurchaseHistory.fromMap(Map<String, dynamic> m) => PurchaseHistory(
    supplierId: (m['supplierId'] as int?) ?? 0,
    rate: (m['rate'] as num?)?.toDouble() ?? 0.0,
    quantity: (m['quantity'] as num?)?.toDouble() ?? 0.0,
    date: (m['date'] as String?) ?? '',
    purchaseId: m['purchaseId'] as int?,
    timestamp: (m['timestamp'] as int?) ?? 0,
  );
}

class StockItem {
  final int id;
  String name;
  String category;
  double purchaseRate;
  double quantity;
  double alertLimit;
  int? lastSupplierId;
  List<PurchaseHistory> purchaseHistory;
  final String createdAt;
  String unit;

  StockItem({
    required this.id, required this.name, required this.category,
    required this.purchaseRate, required this.quantity,
    required this.alertLimit, this.lastSupplierId,
    required this.purchaseHistory, required this.createdAt,
    this.unit = 'KG',
  });

  bool get isLowStock => quantity <= alertLimit && alertLimit > 0;

  Map<String, dynamic> toMap() => {
    'id': id, 'name': name, 'category': category,
    'purchaseRate': purchaseRate, 'quantity': quantity,
    'alertLimit': alertLimit, 'lastSupplierId': lastSupplierId,
    'createdAt': createdAt, 'unit': unit,
  };

  factory StockItem.fromMap(Map<String, dynamic> m) => StockItem(
    id: (m['id'] as int?) ?? 0,
    name: (m['name'] as String?) ?? '',
    category: (m['category'] as String?) ?? '',
    purchaseRate: (m['purchaseRate'] as num?)?.toDouble() ?? 0.0,
    quantity: (m['quantity'] as num?)?.toDouble() ?? 0.0,
    alertLimit: (m['alertLimit'] as num?)?.toDouble() ?? 0.0,
    lastSupplierId: m['lastSupplierId'] as int?,
    purchaseHistory: [],
    createdAt: (m['createdAt'] as String?) ?? DateTime.now().toIso8601String(),
    unit: (m['unit'] as String?) ?? 'KG',
  );

  StockItem copyWith({String? name, String? category, double? purchaseRate,
      double? quantity, double? alertLimit, int? lastSupplierId,
      List<PurchaseHistory>? purchaseHistory, String? unit}) =>
    StockItem(
      id: id, name: name ?? this.name, category: category ?? this.category,
      purchaseRate: purchaseRate ?? this.purchaseRate,
      quantity: quantity ?? this.quantity,
      alertLimit: alertLimit ?? this.alertLimit,
      lastSupplierId: lastSupplierId ?? this.lastSupplierId,
      purchaseHistory: purchaseHistory ?? this.purchaseHistory,
      createdAt: createdAt, unit: unit ?? this.unit,
    );
}

class SaleItem {
  final int itemId;
  final double qty;
  final double rate;
  final double total;
  final double costRate;

  const SaleItem({
    required this.itemId, required this.qty, required this.rate,
    required this.total, required this.costRate,
  });

  double get profit => total - (costRate * qty);

  Map<String, dynamic> toMap() => {
    'itemId': itemId, 'qty': qty, 'rate': rate,
    'total': total, 'costRate': costRate,
  };

  factory SaleItem.fromMap(Map<String, dynamic> m) => SaleItem(
    itemId: (m['itemId'] as int?) ?? 0,
    qty: (m['qty'] as num?)?.toDouble() ?? 0.0,
    rate: (m['rate'] as num?)?.toDouble() ?? 0.0,
    total: (m['total'] as num?)?.toDouble() ?? 0.0,
    costRate: (m['costRate'] as num?)?.toDouble() ?? 0.0,
  );
}

class Sale {
  final int id;
  final int customerId;
  List<SaleItem> items;
  final double total;
  final double discount;
  final double fee;
  final double cashReceived;
  final double creditAmount;
  final String date;
  final int timestamp;
  final String note;

  Sale({
    required this.id, required this.customerId, required this.items,
    required this.total, required this.discount, required this.fee,
    required this.cashReceived, required this.creditAmount,
    required this.date, required this.timestamp, this.note = '',
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'customerId': customerId, 'total': total,
    'discount': discount, 'fee': fee, 'cashReceived': cashReceived,
    'creditAmount': creditAmount, 'date': date,
    'timestamp': timestamp, 'note': note,
  };

  factory Sale.fromMap(Map<String, dynamic> m) => Sale(
    id: (m['id'] as int?) ?? 0,
    customerId: (m['customerId'] as int?) ?? 0,
    items: [],
    total: (m['total'] as num?)?.toDouble() ?? 0.0,
    discount: (m['discount'] as num?)?.toDouble() ?? 0.0,
    fee: (m['fee'] as num?)?.toDouble() ?? 0.0,
    cashReceived: (m['cashReceived'] as num?)?.toDouble() ?? 0.0,
    creditAmount: (m['creditAmount'] as num?)?.toDouble() ?? 0.0,
    date: (m['date'] as String?) ?? '',
    timestamp: (m['timestamp'] as int?) ?? 0,
    note: (m['note'] as String?) ?? '',
  );
}

class PurchaseItem {
  final int itemId;
  final double qty;
  final double rate;
  final double total;

  const PurchaseItem({
    required this.itemId, required this.qty,
    required this.rate, required this.total,
  });

  Map<String, dynamic> toMap() => {
    'itemId': itemId, 'qty': qty, 'rate': rate, 'total': total,
  };

  factory PurchaseItem.fromMap(Map<String, dynamic> m) => PurchaseItem(
    itemId: (m['itemId'] as int?) ?? 0,
    qty: (m['qty'] as num?)?.toDouble() ?? 0.0,
    rate: (m['rate'] as num?)?.toDouble() ?? 0.0,
    total: (m['total'] as num?)?.toDouble() ?? 0.0,
  );
}

class Purchase {
  final int id;
  final int supplierId;
  List<PurchaseItem> items;
  final double total;
  final double discount;
  final double fee;
  final double cashPaid;
  final double creditAmount;
  final String date;
  final int timestamp;
  final String note;

  Purchase({
    required this.id, required this.supplierId, required this.items,
    required this.total, required this.discount, required this.fee,
    required this.cashPaid, required this.creditAmount,
    required this.date, required this.timestamp, this.note = '',
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'supplierId': supplierId, 'total': total,
    'discount': discount, 'fee': fee, 'cashPaid': cashPaid,
    'creditAmount': creditAmount, 'date': date,
    'timestamp': timestamp, 'note': note,
  };

  factory Purchase.fromMap(Map<String, dynamic> m) => Purchase(
    id: (m['id'] as int?) ?? 0,
    supplierId: (m['supplierId'] as int?) ?? 0,
    items: [],
    total: (m['total'] as num?)?.toDouble() ?? 0.0,
    discount: (m['discount'] as num?)?.toDouble() ?? 0.0,
    fee: (m['fee'] as num?)?.toDouble() ?? 0.0,
    cashPaid: (m['cashPaid'] as num?)?.toDouble() ?? 0.0,
    creditAmount: (m['creditAmount'] as num?)?.toDouble() ?? 0.0,
    date: (m['date'] as String?) ?? '',
    timestamp: (m['timestamp'] as int?) ?? 0,
    note: (m['note'] as String?) ?? '',
  );
}

class Expense {
  final int id;
  String category;
  double amount;
  String note;
  final String date;
  final int timestamp;
  final int? linkedSupplierId;

  Expense({
    required this.id, required this.category, required this.amount,
    this.note = '', required this.date, required this.timestamp,
    this.linkedSupplierId,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'category': category, 'amount': amount,
    'note': note, 'date': date, 'timestamp': timestamp,
    'linkedSupplierId': linkedSupplierId,
  };

  factory Expense.fromMap(Map<String, dynamic> m) => Expense(
    id: (m['id'] as int?) ?? 0,
    category: (m['category'] as String?) ?? '',
    amount: (m['amount'] as num?)?.toDouble() ?? 0.0,
    note: (m['note'] as String?) ?? '',
    date: (m['date'] as String?) ?? '',
    timestamp: (m['timestamp'] as int?) ?? 0,
    linkedSupplierId: m['linkedSupplierId'] as int?,
  );
}

class CashInOutEntry {
  final int id;
  final String type;
  final double amount;
  final String note;
  final String date;
  final int timestamp;

  const CashInOutEntry({
    required this.id, required this.type, required this.amount,
    this.note = '', required this.date, required this.timestamp,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'type': type, 'amount': amount,
    'note': note, 'date': date, 'timestamp': timestamp,
  };

  factory CashInOutEntry.fromMap(Map<String, dynamic> m) => CashInOutEntry(
    id: (m['id'] as int?) ?? 0,
    type: (m['type'] as String?) ?? '',
    amount: (m['amount'] as num?)?.toDouble() ?? 0.0,
    note: (m['note'] as String?) ?? '',
    date: (m['date'] as String?) ?? '',
    timestamp: (m['timestamp'] as int?) ?? 0,
  );
}

class CashLedgerEntry {
  final int id;
  final String date;
  final String type;
  final double amount;
  final double balance;
  final String note;

  const CashLedgerEntry({
    required this.id, required this.date, required this.type,
    required this.amount, required this.balance, this.note = '',
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'date': date, 'type': type,
    'amount': amount, 'balance': balance, 'note': note,
  };

  factory CashLedgerEntry.fromMap(Map<String, dynamic> m) => CashLedgerEntry(
    id: (m['id'] as int?) ?? 0,
    date: (m['date'] as String?) ?? '',
    type: (m['type'] as String?) ?? '',
    amount: (m['amount'] as num?)?.toDouble() ?? 0.0,
    balance: (m['balance'] as num?)?.toDouble() ?? 0.0,
    note: (m['note'] as String?) ?? '',
  );
}

class StockAdjustment {
  final int id;
  final int itemId;
  final String itemName;
  final String adjType;
  final double qty;
  final double rate;
  final double amount;
  final String unit;
  final String date;
  final String note;
  final int timestamp;

  const StockAdjustment({
    required this.id, required this.itemId, required this.itemName,
    required this.adjType, required this.qty, required this.rate,
    required this.amount, required this.unit, required this.date,
    this.note = '', required this.timestamp,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'itemId': itemId, 'itemName': itemName,
    'adjType': adjType, 'qty': qty, 'rate': rate,
    'amount': amount, 'unit': unit, 'date': date,
    'note': note, 'timestamp': timestamp,
  };

  factory StockAdjustment.fromMap(Map<String, dynamic> m) => StockAdjustment(
    id: (m['id'] as int?) ?? 0,
    itemId: (m['itemId'] as int?) ?? 0,
    itemName: (m['itemName'] as String?) ?? '',
    adjType: (m['adjType'] as String?) ?? 'loss',
    qty: (m['qty'] as num?)?.toDouble() ?? 0.0,
    rate: (m['rate'] as num?)?.toDouble() ?? 0.0,
    amount: (m['amount'] as num?)?.toDouble() ?? 0.0,
    unit: (m['unit'] as String?) ?? 'KG',
    date: (m['date'] as String?) ?? '',
    note: (m['note'] as String?) ?? '',
    timestamp: (m['timestamp'] as int?) ?? 0,
  );
}

class Partner {
  final int id;
  String name;
  double sharePercent;
  double totalWithdrawal;
  List<Map<String, dynamic>> withdrawals;

  Partner({
    required this.id, required this.name, required this.sharePercent,
    required this.totalWithdrawal, required this.withdrawals,
  });

  Map<String, dynamic> toMap() => {
    'id': id, 'name': name,
    'sharePercent': sharePercent, 'totalWithdrawal': totalWithdrawal,
  };

  factory Partner.fromMap(Map<String, dynamic> m) => Partner(
    id: (m['id'] as int?) ?? 0,
    name: (m['name'] as String?) ?? '',
    sharePercent: (m['sharePercent'] as num?)?.toDouble() ?? 0.0,
    totalWithdrawal: (m['totalWithdrawal'] as num?)?.toDouble() ?? 0.0,
    withdrawals: [],
  );

  Partner copyWith({String? name, double? sharePercent,
      double? totalWithdrawal, List<Map<String, dynamic>>? withdrawals}) =>
    Partner(id: id, name: name ?? this.name,
      sharePercent: sharePercent ?? this.sharePercent,
      totalWithdrawal: totalWithdrawal ?? this.totalWithdrawal,
      withdrawals: withdrawals ?? this.withdrawals);
}

class UndoAction {
  final String type;
  final Map<String, dynamic> data;
  const UndoAction({required this.type, required this.data});
}


// ============================================================
// UTILS
// ============================================================

class CurrencyUtils {
  static final NumberFormat _fmt = NumberFormat('#,##0.00', 'en_US');
  static String format(double amount) => 'Rs. ${_fmt.format(amount)}';
  static String formatShort(double amount) {
    if (amount >= 1000000) return 'Rs. ${(amount / 1000000).toStringAsFixed(2)}M';
    if (amount >= 1000)    return 'Rs. ${(amount / 1000).toStringAsFixed(1)}K';
    return 'Rs. ${amount.toStringAsFixed(2)}';
  }
}

class AppDateUtils {
  static String today() => DateTime.now().toIso8601String().split('T')[0];
  static String todayFormatted() => DateFormat('dd MMM yyyy').format(DateTime.now());
  static String format(String date) {
    try { return DateFormat('dd MMM yyyy').format(DateTime.parse(date)); }
    catch (_) { return date; }
  }
  static String formatDateTime(int ts) {
    try { return DateFormat('dd MMM yyyy — hh:mm a').format(DateTime.fromMillisecondsSinceEpoch(ts)); }
    catch (_) { return ''; }
  }
}

class AppHelpers {
  static String sanitizeInput(String s) =>
      s.replaceAll(RegExp(r'[<>"\'\\]'), '').trim();
}

// ============================================================
// EXCEPTIONS
// ============================================================

class AppException implements Exception {
  final String message;
  const AppException(this.message);
  @override String toString() => message;
}
class InsufficientStockException extends AppException {
  const InsufficientStockException(String msg) : super(msg);
}
class DatabaseException extends AppException {
  const DatabaseException(String msg) : super(msg);
}
class BackupException extends AppException {
  const BackupException(String msg) : super(msg);
}


// ============================================================
// IQBAL TRADERS — SQLite Database Helper (Clean, Null-Safe)
// ============================================================

class DatabaseHelper {
  static final DatabaseHelper _instance = DatabaseHelper._internal();
  static Database? _db;
  factory DatabaseHelper() => _instance;
  DatabaseHelper._internal();

  Future<Database> get database async {
    _db ??= await _init();
    return _db!;
  }

  Future<Database> _init() async {
    final path = dbPath.join(await getDatabasesPath(), 'iqbal_traders.db');
    return openDatabase(path, version: 1, onCreate: _create);
  }

  Future<void> _create(Database db, int v) async {
    final batch = db.batch();
    batch.execute('''CREATE TABLE settings(key TEXT PRIMARY KEY, value TEXT NOT NULL)''');
    batch.execute('''CREATE TABLE customers(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
      phone TEXT DEFAULT '', address TEXT DEFAULT '',
      balance REAL DEFAULT 0, createdAt TEXT NOT NULL)''');
    batch.execute('''CREATE TABLE customer_transactions(
      id INTEGER PRIMARY KEY AUTOINCREMENT, customerId INTEGER NOT NULL,
      date TEXT NOT NULL, type TEXT NOT NULL, amount REAL NOT NULL,
      note TEXT DEFAULT '', timestamp INTEGER NOT NULL,
      saleId INTEGER, purchaseId INTEGER)''');
    batch.execute('''CREATE TABLE suppliers(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
      phone TEXT DEFAULT '', address TEXT DEFAULT '',
      balance REAL DEFAULT 0, createdAt TEXT NOT NULL)''');
    batch.execute('''CREATE TABLE supplier_transactions(
      id INTEGER PRIMARY KEY AUTOINCREMENT, supplierId INTEGER NOT NULL,
      date TEXT NOT NULL, type TEXT NOT NULL, amount REAL NOT NULL,
      note TEXT DEFAULT '', timestamp INTEGER NOT NULL,
      saleId INTEGER, purchaseId INTEGER)''');
    batch.execute('''CREATE TABLE udhar(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
      phone TEXT DEFAULT '', address TEXT DEFAULT '',
      balance REAL DEFAULT 0, createdAt TEXT NOT NULL)''');
    batch.execute('''CREATE TABLE udhar_transactions(
      id INTEGER PRIMARY KEY AUTOINCREMENT, udharId INTEGER NOT NULL,
      date TEXT NOT NULL, type TEXT NOT NULL, amount REAL NOT NULL,
      note TEXT DEFAULT '', timestamp INTEGER NOT NULL)''');
    batch.execute('''CREATE TABLE stock(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
      category TEXT DEFAULT '', purchaseRate REAL DEFAULT 0,
      quantity REAL DEFAULT 0, alertLimit REAL DEFAULT 0,
      lastSupplierId INTEGER, createdAt TEXT NOT NULL, unit TEXT DEFAULT 'KG')''');
    batch.execute('''CREATE TABLE purchase_history(
      id INTEGER PRIMARY KEY AUTOINCREMENT, itemId INTEGER NOT NULL,
      supplierId INTEGER NOT NULL, rate REAL NOT NULL,
      quantity REAL NOT NULL, date TEXT NOT NULL,
      purchaseId INTEGER, timestamp INTEGER NOT NULL)''');
    batch.execute('''CREATE TABLE stock_adjustments(
      id INTEGER PRIMARY KEY AUTOINCREMENT, itemId INTEGER NOT NULL,
      itemName TEXT NOT NULL, adjType TEXT NOT NULL,
      qty REAL NOT NULL, rate REAL NOT NULL, amount REAL NOT NULL,
      unit TEXT DEFAULT 'KG', date TEXT NOT NULL,
      note TEXT DEFAULT '', timestamp INTEGER NOT NULL)''');
    batch.execute('''CREATE TABLE sales(
      id INTEGER PRIMARY KEY AUTOINCREMENT, customerId INTEGER NOT NULL,
      total REAL NOT NULL, discount REAL DEFAULT 0, fee REAL DEFAULT 0,
      cashReceived REAL DEFAULT 0, creditAmount REAL DEFAULT 0,
      date TEXT NOT NULL, timestamp INTEGER NOT NULL, note TEXT DEFAULT '')''');
    batch.execute('''CREATE TABLE sale_items(
      id INTEGER PRIMARY KEY AUTOINCREMENT, saleId INTEGER NOT NULL,
      itemId INTEGER NOT NULL, qty REAL NOT NULL, rate REAL NOT NULL,
      total REAL NOT NULL, costRate REAL NOT NULL)''');
    batch.execute('''CREATE TABLE purchases(
      id INTEGER PRIMARY KEY AUTOINCREMENT, supplierId INTEGER NOT NULL,
      total REAL NOT NULL, discount REAL DEFAULT 0, fee REAL DEFAULT 0,
      cashPaid REAL DEFAULT 0, creditAmount REAL DEFAULT 0,
      date TEXT NOT NULL, timestamp INTEGER NOT NULL, note TEXT DEFAULT '')''');
    batch.execute('''CREATE TABLE purchase_items(
      id INTEGER PRIMARY KEY AUTOINCREMENT, purchaseId INTEGER NOT NULL,
      itemId INTEGER NOT NULL, qty REAL NOT NULL, rate REAL NOT NULL,
      total REAL NOT NULL)''');
    batch.execute('''CREATE TABLE expenses(
      id INTEGER PRIMARY KEY AUTOINCREMENT, category TEXT NOT NULL,
      amount REAL NOT NULL, note TEXT DEFAULT '',
      date TEXT NOT NULL, timestamp INTEGER NOT NULL,
      linkedSupplierId INTEGER)''');
    batch.execute('''CREATE TABLE cash_in_out(
      id INTEGER PRIMARY KEY AUTOINCREMENT, type TEXT NOT NULL,
      amount REAL NOT NULL, note TEXT DEFAULT '',
      date TEXT NOT NULL, timestamp INTEGER NOT NULL)''');
    batch.execute('''CREATE TABLE cash_ledger(
      id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL,
      type TEXT NOT NULL, amount REAL NOT NULL,
      balance REAL NOT NULL, note TEXT DEFAULT '')''');
    batch.execute('''CREATE TABLE partners(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
      sharePercent REAL NOT NULL, totalWithdrawal REAL DEFAULT 0)''');
    batch.execute('''CREATE TABLE partner_withdrawals(
      id INTEGER PRIMARY KEY AUTOINCREMENT, partnerId INTEGER NOT NULL,
      amount REAL NOT NULL, note TEXT DEFAULT '',
      date TEXT NOT NULL, timestamp INTEGER NOT NULL)''');
    await batch.commit(noResult: true);
    // defaults
    await db.insert('settings', {'key': 'shopName', 'value': 'IQBAL TRADERS'});
    await db.insert('settings', {'key': 'password', 'value': '1234'});
    await db.insert('settings', {'key': 'cashInHand', 'value': '0.0'});
  }

  // ── Settings ────────────────────────────────────────────────
  Future<String> getSetting(String key, {String defaultValue = ''}) async {
    final db = await database;
    final r = await db.query('settings', where: 'key=?', whereArgs: [key]);
    return r.isEmpty ? defaultValue : (r.first['value'] as String);
  }
  Future<void> setSetting(String key, String value) async {
    final db = await database;
    await db.insert('settings', {'key': key, 'value': value},
        conflictAlgorithm: ConflictAlgorithm.replace);
  }
  Future<double> getCashInHand() async =>
      double.tryParse(await getSetting('cashInHand', defaultValue: '0.0')) ?? 0.0;
  Future<void> setCashInHand(double v) async => setSetting('cashInHand', v.toString());

  // ── Customers ───────────────────────────────────────────────
  Future<int> insertCustomer(Customer c) async {
    final db = await database;
    final m = c.toMap()..remove('id');
    return db.insert('customers', m);
  }
  Future<void> updateCustomer(Customer c) async {
    final db = await database;
    await db.update('customers', c.toMap(), where: 'id=?', whereArgs: [c.id]);
  }
  Future<void> deleteCustomer(int id) async {
    final db = await database;
    await db.delete('customers', where: 'id=?', whereArgs: [id]);
    await db.delete('customer_transactions', where: 'customerId=?', whereArgs: [id]);
  }
  Future<List<Customer>> getAllCustomers() async {
    final db = await database;
    final rows = await db.query('customers', orderBy: 'name ASC');
    final list = rows.map(Customer.fromMap).toList();
    for (final c in list) c.transactions = await getCustomerTx(c.id);
    return list;
  }
  Future<void> insertCustomerTx(int cid, AppTransaction t) async {
    final db = await database;
    await db.insert('customer_transactions', {
      'customerId': cid, 'date': t.date, 'type': t.type,
      'amount': t.amount, 'note': t.note, 'timestamp': t.timestamp,
      'saleId': t.saleId, 'purchaseId': t.purchaseId,
    });
  }
  Future<List<AppTransaction>> getCustomerTx(int cid) async {
    final db = await database;
    final rows = await db.query('customer_transactions',
        where: 'customerId=?', whereArgs: [cid], orderBy: 'timestamp ASC');
    return rows.map(AppTransaction.fromMap).toList();
  }
  Future<void> deleteCustomerTxBySaleId(int cid, int saleId) async {
    final db = await database;
    await db.delete('customer_transactions',
        where: 'customerId=? AND saleId=?', whereArgs: [cid, saleId]);
  }
  Future<void> deleteLastCustomerTx(int cid) async {
    final db = await database;
    final r = await db.query('customer_transactions',
        where: 'customerId=?', whereArgs: [cid],
        orderBy: 'timestamp DESC', limit: 1);
    if (r.isNotEmpty) {
      await db.delete('customer_transactions', where: 'id=?', whereArgs: [r.first['id']]);
    }
  }

  // ── Suppliers ───────────────────────────────────────────────
  Future<int> insertSupplier(Supplier s) async {
    final db = await database;
    final m = s.toMap()..remove('id');
    return db.insert('suppliers', m);
  }
  Future<void> updateSupplier(Supplier s) async {
    final db = await database;
    await db.update('suppliers', s.toMap(), where: 'id=?', whereArgs: [s.id]);
  }
  Future<void> deleteSupplier(int id) async {
    final db = await database;
    await db.delete('suppliers', where: 'id=?', whereArgs: [id]);
    await db.delete('supplier_transactions', where: 'supplierId=?', whereArgs: [id]);
  }
  Future<List<Supplier>> getAllSuppliers() async {
    final db = await database;
    final rows = await db.query('suppliers', orderBy: 'name ASC');
    final list = rows.map(Supplier.fromMap).toList();
    for (final s in list) s.transactions = await getSupplierTx(s.id);
    return list;
  }
  Future<void> insertSupplierTx(int sid, AppTransaction t) async {
    final db = await database;
    await db.insert('supplier_transactions', {
      'supplierId': sid, 'date': t.date, 'type': t.type,
      'amount': t.amount, 'note': t.note, 'timestamp': t.timestamp,
      'saleId': t.saleId, 'purchaseId': t.purchaseId,
    });
  }
  Future<List<AppTransaction>> getSupplierTx(int sid) async {
    final db = await database;
    final rows = await db.query('supplier_transactions',
        where: 'supplierId=?', whereArgs: [sid], orderBy: 'timestamp ASC');
    return rows.map(AppTransaction.fromMap).toList();
  }
  Future<void> deleteSupplierTxByPurchaseId(int sid, int pid) async {
    final db = await database;
    await db.delete('supplier_transactions',
        where: 'supplierId=? AND purchaseId=?', whereArgs: [sid, pid]);
  }
  Future<void> deleteLastSupplierTx(int sid) async {
    final db = await database;
    final r = await db.query('supplier_transactions',
        where: 'supplierId=?', whereArgs: [sid],
        orderBy: 'timestamp DESC', limit: 1);
    if (r.isNotEmpty) {
      await db.delete('supplier_transactions', where: 'id=?', whereArgs: [r.first['id']]);
    }
  }

  // ── Udhar ───────────────────────────────────────────────────
  Future<int> insertUdharPerson(UdharPerson u) async {
    final db = await database;
    final m = u.toMap()..remove('id');
    return db.insert('udhar', m);
  }
  Future<void> updateUdharPerson(UdharPerson u) async {
    final db = await database;
    await db.update('udhar', u.toMap(), where: 'id=?', whereArgs: [u.id]);
  }
  Future<void> deleteUdharPerson(int id) async {
    final db = await database;
    await db.delete('udhar', where: 'id=?', whereArgs: [id]);
    await db.delete('udhar_transactions', where: 'udharId=?', whereArgs: [id]);
  }
  Future<List<UdharPerson>> getAllUdharPersons() async {
    final db = await database;
    final rows = await db.query('udhar', orderBy: 'name ASC');
    final list = rows.map(UdharPerson.fromMap).toList();
    for (final u in list) u.transactions = await getUdharTx(u.id);
    return list;
  }
  Future<void> insertUdharTx(int uid, AppTransaction t) async {
    final db = await database;
    await db.insert('udhar_transactions', {
      'udharId': uid, 'date': t.date, 'type': t.type,
      'amount': t.amount, 'note': t.note, 'timestamp': t.timestamp,
    });
  }
  Future<List<AppTransaction>> getUdharTx(int uid) async {
    final db = await database;
    final rows = await db.query('udhar_transactions',
        where: 'udharId=?', whereArgs: [uid], orderBy: 'timestamp ASC');
    return rows.map(AppTransaction.fromMap).toList();
  }
  Future<void> deleteLastUdharTx(int uid) async {
    final db = await database;
    final r = await db.query('udhar_transactions',
        where: 'udharId=?', whereArgs: [uid],
        orderBy: 'timestamp DESC', limit: 1);
    if (r.isNotEmpty) {
      await db.delete('udhar_transactions', where: 'id=?', whereArgs: [r.first['id']]);
    }
  }

  // ── Stock ───────────────────────────────────────────────────
  Future<int> insertStockItem(StockItem item) async {
    final db = await database;
    final m = item.toMap()..remove('id');
    return db.insert('stock', m);
  }
  Future<void> updateStockItem(StockItem item) async {
    final db = await database;
    await db.update('stock', item.toMap(), where: 'id=?', whereArgs: [item.id]);
  }
  Future<void> deleteStockItem(int id) async {
    final db = await database;
    await db.delete('stock', where: 'id=?', whereArgs: [id]);
    await db.delete('purchase_history', where: 'itemId=?', whereArgs: [id]);
    await db.delete('stock_adjustments', where: 'itemId=?', whereArgs: [id]);
  }
  Future<List<StockItem>> getAllStockItems() async {
    final db = await database;
    final rows = await db.query('stock', orderBy: 'name ASC');
    final list = rows.map(StockItem.fromMap).toList();
    for (final i in list) i.purchaseHistory = await getPurchaseHistory(i.id);
    return list;
  }
  Future<void> insertPurchaseHistory(int itemId, PurchaseHistory ph) async {
    final db = await database;
    await db.insert('purchase_history', {
      'itemId': itemId, 'supplierId': ph.supplierId, 'rate': ph.rate,
      'quantity': ph.quantity, 'date': ph.date,
      'purchaseId': ph.purchaseId, 'timestamp': ph.timestamp,
    });
  }
  Future<List<PurchaseHistory>> getPurchaseHistory(int itemId) async {
    final db = await database;
    final rows = await db.query('purchase_history',
        where: 'itemId=?', whereArgs: [itemId], orderBy: 'timestamp ASC');
    return rows.map(PurchaseHistory.fromMap).toList();
  }
  Future<void> deleteLastPurchaseHistory(int itemId) async {
    final db = await database;
    final r = await db.query('purchase_history',
        where: 'itemId=?', whereArgs: [itemId],
        orderBy: 'timestamp DESC', limit: 1);
    if (r.isNotEmpty) {
      await db.delete('purchase_history', where: 'id=?', whereArgs: [r.first['id']]);
    }
  }
  Future<int> insertStockAdjustment(StockAdjustment adj) async {
    final db = await database;
    final m = adj.toMap()..remove('id');
    return db.insert('stock_adjustments', m);
  }
  Future<void> deleteStockAdjustment(int id) async {
    final db = await database;
    await db.delete('stock_adjustments', where: 'id=?', whereArgs: [id]);
  }
  Future<List<StockAdjustment>> getAllStockAdjustments() async {
    final db = await database;
    final rows = await db.query('stock_adjustments', orderBy: 'timestamp DESC');
    return rows.map(StockAdjustment.fromMap).toList();
  }

  // ── Sales ───────────────────────────────────────────────────
  Future<int> insertSale(Sale sale) async {
    final db = await database;
    return db.transaction((txn) async {
      final m = sale.toMap()..remove('id');
      final id = await txn.insert('sales', m);
      for (final item in sale.items) {
        await txn.insert('sale_items', {...item.toMap(), 'saleId': id});
      }
      return id;
    });
  }
  Future<void> deleteSale(int id) async {
    final db = await database;
    await db.transaction((txn) async {
      await txn.delete('sales', where: 'id=?', whereArgs: [id]);
      await txn.delete('sale_items', where: 'saleId=?', whereArgs: [id]);
    });
  }
  Future<List<Sale>> getAllSales() async {
    final db = await database;
    final rows = await db.query('sales', orderBy: 'timestamp DESC');
    final list = rows.map(Sale.fromMap).toList();
    for (final s in list) s.items = await _getSaleItems(db, s.id);
    return list;
  }
  Future<List<SaleItem>> _getSaleItems(Database db, int sid) async {
    final rows = await db.query('sale_items', where: 'saleId=?', whereArgs: [sid]);
    return rows.map(SaleItem.fromMap).toList();
  }

  // ── Purchases ───────────────────────────────────────────────
  Future<int> insertPurchase(Purchase p) async {
    final db = await database;
    return db.transaction((txn) async {
      final m = p.toMap()..remove('id');
      final id = await txn.insert('purchases', m);
      for (final item in p.items) {
        await txn.insert('purchase_items', {...item.toMap(), 'purchaseId': id});
      }
      return id;
    });
  }
  Future<void> deletePurchase(int id) async {
    final db = await database;
    await db.transaction((txn) async {
      await txn.delete('purchases', where: 'id=?', whereArgs: [id]);
      await txn.delete('purchase_items', where: 'purchaseId=?', whereArgs: [id]);
    });
  }
  Future<List<Purchase>> getAllPurchases() async {
    final db = await database;
    final rows = await db.query('purchases', orderBy: 'timestamp DESC');
    final list = rows.map(Purchase.fromMap).toList();
    for (final p in list) p.items = await _getPurchaseItems(db, p.id);
    return list;
  }
  Future<List<PurchaseItem>> _getPurchaseItems(Database db, int pid) async {
    final rows = await db.query('purchase_items', where: 'purchaseId=?', whereArgs: [pid]);
    return rows.map(PurchaseItem.fromMap).toList();
  }

  // ── Expenses ────────────────────────────────────────────────
  Future<int> insertExpense(Expense e) async {
    final db = await database;
    final m = e.toMap()..remove('id');
    return db.insert('expenses', m);
  }
  Future<void> updateExpense(Expense e) async {
    final db = await database;
    await db.update('expenses', e.toMap(), where: 'id=?', whereArgs: [e.id]);
  }
  Future<void> deleteExpense(int id) async {
    final db = await database;
    await db.delete('expenses', where: 'id=?', whereArgs: [id]);
  }
  Future<List<Expense>> getAllExpenses() async {
    final db = await database;
    final rows = await db.query('expenses', orderBy: 'timestamp DESC');
    return rows.map(Expense.fromMap).toList();
  }

  // ── Cash In/Out ─────────────────────────────────────────────
  Future<int> insertCashInOut(CashInOutEntry e) async {
    final db = await database;
    final m = e.toMap()..remove('id');
    return db.insert('cash_in_out', m);
  }
  Future<void> deleteCashInOut(int id) async {
    final db = await database;
    await db.delete('cash_in_out', where: 'id=?', whereArgs: [id]);
  }
  Future<List<CashInOutEntry>> getAllCashInOut() async {
    final db = await database;
    final rows = await db.query('cash_in_out', orderBy: 'timestamp DESC');
    return rows.map(CashInOutEntry.fromMap).toList();
  }
  Future<int> insertCashLedger(CashLedgerEntry e) async {
    final db = await database;
    final m = e.toMap()..remove('id');
    return db.insert('cash_ledger', m);
  }
  Future<void> deleteLastCashLedger() async {
    final db = await database;
    final r = await db.query('cash_ledger', orderBy: 'id DESC', limit: 1);
    if (r.isNotEmpty) {
      await db.delete('cash_ledger', where: 'id=?', whereArgs: [r.first['id']]);
    }
  }
  Future<List<CashLedgerEntry>> getCashLedger() async {
    final db = await database;
    final rows = await db.query('cash_ledger', orderBy: 'id DESC');
    return rows.map(CashLedgerEntry.fromMap).toList();
  }

  // ── Partners ────────────────────────────────────────────────
  Future<int> insertPartner(Partner p) async {
    final db = await database;
    final m = p.toMap()..remove('id');
    return db.insert('partners', m);
  }
  Future<void> updatePartner(Partner p) async {
    final db = await database;
    await db.update('partners', p.toMap(), where: 'id=?', whereArgs: [p.id]);
  }
  Future<void> deletePartner(int id) async {
    final db = await database;
    await db.delete('partners', where: 'id=?', whereArgs: [id]);
    await db.delete('partner_withdrawals', where: 'partnerId=?', whereArgs: [id]);
  }
  Future<List<Partner>> getAllPartners() async {
    final db = await database;
    final rows = await db.query('partners');
    final list = rows.map(Partner.fromMap).toList();
    for (final p in list) {
      final wRows = await db.query('partner_withdrawals',
          where: 'partnerId=?', whereArgs: [p.id], orderBy: 'timestamp DESC');
      p.withdrawals = wRows.map((r) => {
        'amount': r['amount'], 'note': r['note'],
        'date': r['date'], 'timestamp': r['timestamp'],
      }).toList();
    }
    return list;
  }
  Future<void> insertPartnerWithdrawal(
      int pid, double amount, String note, String date, int ts) async {
    final db = await database;
    await db.insert('partner_withdrawals', {
      'partnerId': pid, 'amount': amount,
      'note': note, 'date': date, 'timestamp': ts,
    });
  }

  // ── Backup / Restore ────────────────────────────────────────
  Future<Map<String, dynamic>> exportAll() async {
    return {
      'shopName': await getSetting('shopName', defaultValue: 'IQBAL TRADERS'),
      'password': await getSetting('password', defaultValue: '1234'),
      'cashInHand': {
        'currentBalance': await getCashInHand(),
        'ledger': (await getCashLedger()).map((h) => h.toMap()).toList(),
      },
      'customers':   (await getAllCustomers()).map((c) => {...c.toMap(), 'transactions': c.transactions.map((t) => t.toMap()).toList()}).toList(),
      'suppliers':   (await getAllSuppliers()).map((s) => {...s.toMap(), 'transactions': s.transactions.map((t) => t.toMap()).toList()}).toList(),
      'udhar':       (await getAllUdharPersons()).map((u) => {...u.toMap(), 'transactions': u.transactions.map((t) => t.toMap()).toList()}).toList(),
      'stock':       (await getAllStockItems()).map((i) => {...i.toMap(), 'purchaseHistory': i.purchaseHistory.map((h) => h.toMap()).toList()}).toList(),
      'sales':       (await getAllSales()).map((s) => {...s.toMap(), 'items': s.items.map((i) => i.toMap()).toList()}).toList(),
      'purchases':   (await getAllPurchases()).map((p) => {...p.toMap(), 'items': p.items.map((i) => i.toMap()).toList()}).toList(),
      'expenses':    (await getAllExpenses()).map((e) => e.toMap()).toList(),
      'cashInOut':   (await getAllCashInOut()).map((c) => c.toMap()).toList(),
      'stockAdj':    (await getAllStockAdjustments()).map((a) => a.toMap()).toList(),
      'partners':    (await getAllPartners()).map((p) => {...p.toMap(), 'withdrawals': p.withdrawals}).toList(),
      'exportedAt':  DateTime.now().toIso8601String(),
      'version':     '2.0',
    };
  }

  Future<void> importAll(Map<String, dynamic> data) async {
    final db = await database;
    await db.transaction((txn) async {
      for (final t in [
        'customer_transactions','supplier_transactions','udhar_transactions',
        'purchase_history','stock_adjustments','sale_items','purchase_items',
        'cash_ledger','partner_withdrawals','customers','suppliers','udhar',
        'stock','sales','purchases','expenses','cash_in_out','partners',
      ]) { await txn.delete(t); }

      await txn.insert('settings', {'key':'shopName','value': data['shopName'] ?? 'IQBAL TRADERS'}, conflictAlgorithm: ConflictAlgorithm.replace);
      await txn.insert('settings', {'key':'password','value': data['password'] ?? '1234'}, conflictAlgorithm: ConflictAlgorithm.replace);
      final cash = (data['cashInHand']?['currentBalance'] as num?)?.toDouble() ?? 0.0;
      await txn.insert('settings', {'key':'cashInHand','value': cash.toString()}, conflictAlgorithm: ConflictAlgorithm.replace);

      for (final c in (data['customers'] as List? ?? [])) {
        final id = await txn.insert('customers', {'id':c['id'],'name':c['name']??'','phone':c['phone']??'','address':c['address']??'','balance':(c['balance'] as num?)?.toDouble()??0.0,'createdAt':c['createdAt']??''}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final t in (c['transactions'] as List? ?? [])) {
          await txn.insert('customer_transactions', {'customerId':id,'date':t['date']??'','type':t['type']??'','amount':(t['amount'] as num?)?.toDouble()??0.0,'note':t['note']??'','timestamp':t['timestamp']??0,'saleId':t['saleId'],'purchaseId':t['purchaseId']});
        }
      }
      for (final s in (data['suppliers'] as List? ?? [])) {
        final id = await txn.insert('suppliers', {'id':s['id'],'name':s['name']??'','phone':s['phone']??'','address':s['address']??'','balance':(s['balance'] as num?)?.toDouble()??0.0,'createdAt':s['createdAt']??''}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final t in (s['transactions'] as List? ?? [])) {
          await txn.insert('supplier_transactions', {'supplierId':id,'date':t['date']??'','type':t['type']??'','amount':(t['amount'] as num?)?.toDouble()??0.0,'note':t['note']??'','timestamp':t['timestamp']??0,'saleId':t['saleId'],'purchaseId':t['purchaseId']});
        }
      }
      for (final u in (data['udhar'] as List? ?? [])) {
        final id = await txn.insert('udhar', {'id':u['id'],'name':u['name']??'','phone':u['phone']??'','address':u['address']??'','balance':(u['balance'] as num?)?.toDouble()??0.0,'createdAt':u['createdAt']??''}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final t in (u['transactions'] as List? ?? [])) {
          await txn.insert('udhar_transactions', {'udharId':id,'date':t['date']??'','type':t['type']??'','amount':(t['amount'] as num?)?.toDouble()??0.0,'note':t['note']??'','timestamp':t['timestamp']??0});
        }
      }
      for (final i in (data['stock'] as List? ?? [])) {
        final id = await txn.insert('stock', {'id':i['id'],'name':i['name']??'','category':i['category']??'','purchaseRate':(i['purchaseRate'] as num?)?.toDouble()??0.0,'quantity':(i['quantity'] as num?)?.toDouble()??0.0,'alertLimit':(i['alertLimit'] as num?)?.toDouble()??0.0,'lastSupplierId':i['lastSupplierId'],'createdAt':i['createdAt']??'','unit':i['unit']??'KG'}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final h in (i['purchaseHistory'] as List? ?? [])) {
          await txn.insert('purchase_history', {'itemId':id,'supplierId':h['supplierId']??0,'rate':(h['rate'] as num?)?.toDouble()??0.0,'quantity':(h['quantity'] as num?)?.toDouble()??0.0,'date':h['date']??'','purchaseId':h['purchaseId'],'timestamp':h['timestamp']??0});
        }
      }
      for (final s in (data['sales'] as List? ?? [])) {
        final id = await txn.insert('sales', {'id':s['id'],'customerId':s['customerId']??0,'total':(s['total'] as num?)?.toDouble()??0.0,'discount':(s['discount'] as num?)?.toDouble()??0.0,'fee':(s['fee'] as num?)?.toDouble()??0.0,'cashReceived':(s['cashReceived'] as num?)?.toDouble()??0.0,'creditAmount':(s['creditAmount'] as num?)?.toDouble()??0.0,'date':s['date']??'','timestamp':s['timestamp']??0,'note':s['note']??''}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final item in (s['items'] as List? ?? [])) {
          await txn.insert('sale_items', {'saleId':id,'itemId':item['itemId']??0,'qty':(item['qty'] as num?)?.toDouble()??0.0,'rate':(item['rate'] as num?)?.toDouble()??0.0,'total':(item['total'] as num?)?.toDouble()??0.0,'costRate':(item['costRate'] as num?)?.toDouble()??0.0});
        }
      }
      for (final p in (data['purchases'] as List? ?? [])) {
        final id = await txn.insert('purchases', {'id':p['id'],'supplierId':p['supplierId']??0,'total':(p['total'] as num?)?.toDouble()??0.0,'discount':(p['discount'] as num?)?.toDouble()??0.0,'fee':(p['fee'] as num?)?.toDouble()??0.0,'cashPaid':(p['cashPaid'] as num?)?.toDouble()??0.0,'creditAmount':(p['creditAmount'] as num?)?.toDouble()??0.0,'date':p['date']??'','timestamp':p['timestamp']??0,'note':p['note']??''}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final item in (p['items'] as List? ?? [])) {
          await txn.insert('purchase_items', {'purchaseId':id,'itemId':item['itemId']??0,'qty':(item['qty'] as num?)?.toDouble()??0.0,'rate':(item['rate'] as num?)?.toDouble()??0.0,'total':(item['total'] as num?)?.toDouble()??0.0});
        }
      }
      for (final e in (data['expenses'] as List? ?? [])) {
        await txn.insert('expenses', {'id':e['id'],'category':e['category']??'','amount':(e['amount'] as num?)?.toDouble()??0.0,'note':e['note']??'','date':e['date']??'','timestamp':e['timestamp']??0,'linkedSupplierId':e['linkedSupplierId']}, conflictAlgorithm: ConflictAlgorithm.replace);
      }
      for (final c in (data['cashInOut'] as List? ?? [])) {
        await txn.insert('cash_in_out', {'id':c['id'],'type':c['type']??'','amount':(c['amount'] as num?)?.toDouble()??0.0,'note':c['note']??'','date':c['date']??'','timestamp':c['timestamp']??0}, conflictAlgorithm: ConflictAlgorithm.replace);
      }
      for (final a in (data['stockAdj'] as List? ?? [])) {
        await txn.insert('stock_adjustments', {'id':a['id'],'itemId':a['itemId']??0,'itemName':a['itemName']??'','adjType':a['adjType']??'loss','qty':(a['qty'] as num?)?.toDouble()??0.0,'rate':(a['rate'] as num?)?.toDouble()??0.0,'amount':(a['amount'] as num?)?.toDouble()??0.0,'unit':a['unit']??'KG','date':a['date']??'','note':a['note']??'','timestamp':a['timestamp']??0}, conflictAlgorithm: ConflictAlgorithm.replace);
      }
      for (final pt in (data['partners'] as List? ?? [])) {
        final id = await txn.insert('partners', {'id':pt['id'],'name':pt['name']??'','sharePercent':(pt['sharePercent'] as num?)?.toDouble()??0.0,'totalWithdrawal':(pt['totalWithdrawal'] as num?)?.toDouble()??0.0}, conflictAlgorithm: ConflictAlgorithm.replace);
        for (final w in (pt['withdrawals'] as List? ?? [])) {
          await txn.insert('partner_withdrawals', {'partnerId':id,'amount':(w['amount'] as num?)?.toDouble()??0.0,'note':w['note']??'','date':w['date']??'','timestamp':w['timestamp']??0});
        }
      }
      if (data['cashInHand']?['ledger'] != null) {
        for (final h in (data['cashInHand']['ledger'] as List)) {
          await txn.insert('cash_ledger', {'date':h['date']??'','type':h['type']??'','amount':(h['amount'] as num?)?.toDouble()??0.0,'balance':(h['balance'] as num?)?.toDouble()??0.0,'note':h['note']??''});
        }
      }
    });
  }

  Future<void> clearTransactions() async {
    final db = await database;
    await db.transaction((txn) async {
      for (final t in ['customer_transactions','supplier_transactions','udhar_transactions','sale_items','purchase_items','cash_ledger','stock_adjustments','partner_withdrawals','sales','purchases','expenses','cash_in_out']) {
        await txn.delete(t);
      }
      await txn.rawUpdate('UPDATE customers SET balance=0');
      await txn.rawUpdate('UPDATE suppliers SET balance=0');
      await txn.rawUpdate('UPDATE udhar SET balance=0');
      await txn.rawUpdate('UPDATE partners SET totalWithdrawal=0');
      await txn.insert('settings', {'key':'cashInHand','value':'0.0'}, conflictAlgorithm: ConflictAlgorithm.replace);
    });
  }

  Future<void> fullReset() async {
    final db = await database;
    await db.transaction((txn) async {
      for (final t in ['customer_transactions','supplier_transactions','udhar_transactions','purchase_history','stock_adjustments','sale_items','purchase_items','cash_ledger','partner_withdrawals','customers','suppliers','udhar','stock','sales','purchases','expenses','cash_in_out','partners']) {
        await txn.delete(t);
      }
      await txn.insert('settings', {'key':'shopName','value':'IQBAL TRADERS'}, conflictAlgorithm: ConflictAlgorithm.replace);
      await txn.insert('settings', {'key':'password','value':'1234'}, conflictAlgorithm: ConflictAlgorithm.replace);
      await txn.insert('settings', {'key':'cashInHand','value':'0.0'}, conflictAlgorithm: ConflictAlgorithm.replace);
    });
  }
}


// ============================================================
// IQBAL TRADERS — Business Logic Provider (Clean, Null-Safe)
// ============================================================

class BusinessProvider extends ChangeNotifier {
  final _db = DatabaseHelper();

  String shopName        = 'IQBAL TRADERS';
  double cashInHand      = 0.0;
  List<CashLedgerEntry>  cashLedger     = [];
  List<Customer>         customers      = [];
  List<Supplier>         suppliers      = [];
  List<UdharPerson>      udharPersons   = [];
  List<StockItem>        stockItems     = [];
  List<Sale>             sales          = [];
  List<Purchase>         purchases      = [];
  List<Expense>          expenses       = [];
  List<CashInOutEntry>   cashInOutList  = [];
  List<StockAdjustment>  stockAdj       = [];
  List<Partner>          partners       = [];
  List<UndoAction>       undoStack      = [];
  bool  isLoading   = true;
  String? errorMsg;

  // ── Init ────────────────────────────────────────────────────
  Future<void> initialize() async {
    isLoading = true;
    notifyListeners();
    try { await _reload(); } catch (e) { errorMsg = e.toString(); }
    isLoading = false;
    notifyListeners();
  }

  Future<void> _reload() async {
    shopName     = await _db.getSetting('shopName', defaultValue: 'IQBAL TRADERS');
    cashInHand   = await _db.getCashInHand();
    cashLedger   = await _db.getCashLedger();
    customers    = await _db.getAllCustomers();
    suppliers    = await _db.getAllSuppliers();
    udharPersons = await _db.getAllUdharPersons();
    stockItems   = await _db.getAllStockItems();
    sales        = await _db.getAllSales();
    purchases    = await _db.getAllPurchases();
    expenses     = await _db.getAllExpenses();
    cashInOutList = await _db.getAllCashInOut();
    stockAdj     = await _db.getAllStockAdjustments();
    partners     = await _db.getAllPartners();
  }

  Future<void> reload() async { await _reload(); notifyListeners(); }

  // ── Settings ────────────────────────────────────────────────
  Future<void> saveShopName(String n) async {
    await _db.setSetting('shopName', n); shopName = n; notifyListeners();
  }
  Future<bool> verifyPassword(String p) async =>
      (await _db.getSetting('password', defaultValue: '1234')) == p;
  Future<void> changePassword(String p) => _db.setSetting('password', p);
  Future<void> setOpeningBalance(double amount) async {
    cashInHand = amount;
    await _db.setCashInHand(amount);
    await _addCashLedger(amount, 'ابتدائی بیلنس', _today());
    notifyListeners();
  }

  // ── WAC ─────────────────────────────────────────────────────
  double getWac(int itemId) {
    try { return stockItems.firstWhere((i) => i.id == itemId).purchaseRate; }
    catch (_) { return 0; }
  }

  double _newWac(int itemId, double newQty, double newRate) {
    try {
      final item = stockItems.firstWhere((i) => i.id == itemId, orElse: ()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
      final total = item.quantity * item.purchaseRate + newQty * newRate;
      final qty   = item.quantity + newQty;
      return qty > 0 ? total / qty : newRate;
    } catch (_) { return newRate; }
  }

  // ── Profit ──────────────────────────────────────────────────
  double get totalGrossProfit => sales.fold(0.0, (s, sale) =>
      s + sale.items.fold(0.0, (ss, i) =>
          ss + i.total - (i.costRate > 0 ? i.costRate : getWac(i.itemId)) * i.qty));

  double get totalExpenses    => expenses.fold(0.0, (s, e) => s + e.amount);
  double get totalAdjLoss     => stockAdj.where((a) => a.adjType=='loss').fold(0.0,(s,a)=>s+a.amount);
  double get totalAdjGain     => stockAdj.where((a) => a.adjType=='gain').fold(0.0,(s,a)=>s+a.amount);
  double get totalCustDisc    => customers.expand((c)=>c.transactions).where((t)=>t.type=='ڈسکاؤنٹ').fold(0.0,(s,t)=>s+t.amount.abs());
  double get totalSuppDisc    => suppliers.expand((s)=>s.transactions).where((t)=>t.type=='ڈسکاؤنٹ').fold(0.0,(s,t)=>s+t.amount.abs());
  double get totalCustTax     => customers.expand((c)=>c.transactions).where((t)=>t.type=='ٹیکس').fold(0.0,(s,t)=>s+t.amount.abs());
  double get netProfit        => totalGrossProfit - totalExpenses - totalAdjLoss + totalAdjGain - totalCustDisc + totalSuppDisc + totalCustTax;

  double get totalStockValue  => stockItems.fold(0.0, (s,i)=>s+(i.quantity*i.purchaseRate));
  double get totalReceivables => customers.fold(0.0,(s,c)=>s+(c.balance>0?c.balance:0));
  double get totalPayables    => suppliers.fold(0.0,(s,x)=>s+(x.balance>0?x.balance:0));
  int    get lowStockCount    => stockItems.where((i)=>i.isLowStock).length;

  String _today() => DateTime.now().toIso8601String().split('T')[0];

  // Today
  List<Sale>         get todaySales       => sales.where((s)=>s.date==_today()).toList();
  List<Purchase>     get todayPurchases   => purchases.where((p)=>p.date==_today()).toList();
  List<Expense>      get todayExpList     => expenses.where((e)=>e.date==_today()).toList();
  List<CashInOutEntry> get todayCashList  => cashInOutList.where((c)=>c.date==_today()).toList();
  double get todaySalesTotal    => todaySales.fold(0.0,(s,x)=>s+x.total);
  double get todayCashReceived  => todaySales.fold(0.0,(s,x)=>s+x.cashReceived);
  double get todayPurchTotal    => todayPurchases.fold(0.0,(s,x)=>s+x.total);
  double get todayCashPaid      => todayPurchases.fold(0.0,(s,x)=>s+x.cashPaid);
  double get todayExpTotal      => todayExpList.fold(0.0,(s,x)=>s+x.amount);
  double get todayGrossProfit   => todaySales.fold(0.0,(s,sale)=>
      s+sale.items.fold(0.0,(ss,i)=>ss+i.total-(i.costRate>0?i.costRate:getWac(i.itemId))*i.qty));

  double getMonthlyProfit(int m, int y) {
    double p=0,ex=0;
    for (final s in sales) { final d=DateTime.tryParse(s.date); if(d!=null&&d.month==m&&d.year==y){ for(final i in s.items){p+=i.total-(i.costRate>0?i.costRate:getWac(i.itemId))*i.qty;}}}
    for (final e in expenses){final d=DateTime.tryParse(e.date);if(d!=null&&d.month==m&&d.year==y)ex+=e.amount;}
    return p-ex;
  }

  double calcItemProfit(int itemId) {
    double p=0;
    for(final s in sales){for(final i in s.items.where((x)=>x.itemId==itemId)){p+=i.total-(i.costRate>0?i.costRate:getWac(itemId))*i.qty;}}
    return p;
  }

  // ── Cash Ledger ─────────────────────────────────────────────
  Future<void> _addCashLedger(double change, String type, String date, {String note=''}) async {
    cashInHand += change;
    await _db.setCashInHand(cashInHand);
    final entry = CashLedgerEntry(id:0, date:date, type:type, amount:change, balance:cashInHand, note:note);
    await _db.insertCashLedger(entry);
    cashLedger = await _db.getCashLedger();
  }

  // ── Customers ───────────────────────────────────────────────
  Future<void> addCustomer(Customer c) async {
    final id = await _db.insertCustomer(c);
    customers.add(Customer(id:id,name:c.name,phone:c.phone,address:c.address,balance:0,transactions:[],createdAt:c.createdAt));
    undoStack.add(UndoAction(type:'add_customer',data:{'id':id}));
    notifyListeners();
  }
  Future<void> updateCustomer(Customer c) async {
    await _db.updateCustomer(c);
    final i = customers.indexWhere((x)=>x.id==c.id);
    if (i!=-1) customers[i]=c;
    notifyListeners();
  }
  Future<void> deleteCustomer(int id) async {
    await _db.deleteCustomer(id);
    customers.removeWhere((c)=>c.id==id);
    notifyListeners();
  }
  Future<void> receiveCustomerPayment(int cid, double amount, String date, String note) async {
    final c = customers.firstWhere((x)=>x.id==cid, orElse: ()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    c.balance -= amount;
    await _db.updateCustomer(c);
    final t = AppTransaction(date:date,type:'وصولی',amount:-amount,note:note,timestamp:ts);
    await _db.insertCustomerTx(cid, t);
    c.transactions.add(t);
    await _addCashLedger(amount,'کسٹمر وصولی',date,note:note);
    undoStack.add(UndoAction(type:'cust_receipt',data:{'cid':cid,'amount':amount,'ts':ts}));
    notifyListeners();
  }
  Future<void> addCustomerDiscount(int cid, double amount, String date, String note) async {
    final c = customers.firstWhere((x)=>x.id==cid, orElse: ()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    c.balance -= amount;
    await _db.updateCustomer(c);
    final t = AppTransaction(date:date,type:'ڈسکاؤنٹ',amount:-amount,note:note,timestamp:ts);
    await _db.insertCustomerTx(cid, t);
    c.transactions.add(t);
    undoStack.add(UndoAction(type:'cust_disc',data:{'cid':cid,'amount':amount,'ts':ts}));
    notifyListeners();
  }
  Future<void> addCustomerTax(int cid, double amount, String date, String note) async {
    final c = customers.firstWhere((x)=>x.id==cid, orElse: ()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    c.balance += amount;
    await _db.updateCustomer(c);
    final t = AppTransaction(date:date,type:'ٹیکس',amount:amount,note:note,timestamp:ts);
    await _db.insertCustomerTx(cid, t);
    c.transactions.add(t);
    undoStack.add(UndoAction(type:'cust_tax',data:{'cid':cid,'amount':amount,'ts':ts}));
    notifyListeners();
  }

  // ── Suppliers ───────────────────────────────────────────────
  Future<void> addSupplier(Supplier s) async {
    final id = await _db.insertSupplier(s);
    suppliers.add(Supplier(id:id,name:s.name,phone:s.phone,address:s.address,balance:0,transactions:[],createdAt:s.createdAt));
    undoStack.add(UndoAction(type:'add_supplier',data:{'id':id}));
    notifyListeners();
  }
  Future<void> updateSupplier(Supplier s) async {
    await _db.updateSupplier(s);
    final i = suppliers.indexWhere((x)=>x.id==s.id);
    if (i!=-1) suppliers[i]=s;
    notifyListeners();
  }
  Future<void> deleteSupplier(int id) async {
    await _db.deleteSupplier(id);
    suppliers.removeWhere((s)=>s.id==id);
    notifyListeners();
  }
  Future<void> paySupplier(int sid, double amount, String date, String note) async {
    final s = suppliers.firstWhere((x)=>x.id==sid, orElse: ()=>Supplier(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    s.balance -= amount;
    await _db.updateSupplier(s);
    final t = AppTransaction(date:date,type:'ادائیگی',amount:-amount,note:note,timestamp:ts);
    await _db.insertSupplierTx(sid, t);
    s.transactions.add(t);
    await _addCashLedger(-amount,'سپلائر ادائیگی',date,note:note);
    undoStack.add(UndoAction(type:'supp_pay',data:{'sid':sid,'amount':amount,'ts':ts}));
    notifyListeners();
  }
  Future<void> addSupplierDiscount(int sid, double amount, String date, String note) async {
    final s = suppliers.firstWhere((x)=>x.id==sid, orElse: ()=>Supplier(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    s.balance -= amount;
    await _db.updateSupplier(s);
    final t = AppTransaction(date:date,type:'ڈسکاؤنٹ',amount:-amount,note:note,timestamp:ts);
    await _db.insertSupplierTx(sid, t);
    s.transactions.add(t);
    undoStack.add(UndoAction(type:'supp_disc',data:{'sid':sid,'amount':amount,'ts':ts}));
    notifyListeners();
  }

  // ── Udhar ───────────────────────────────────────────────────
  Future<void> addUdharPerson(UdharPerson u) async {
    final id = await _db.insertUdharPerson(u);
    udharPersons.add(UdharPerson(id:id,name:u.name,phone:u.phone,address:u.address,balance:0,transactions:[],createdAt:u.createdAt));
    notifyListeners();
  }
  Future<void> updateUdharPerson(UdharPerson u) async {
    await _db.updateUdharPerson(u);
    final i = udharPersons.indexWhere((x)=>x.id==u.id);
    if (i!=-1) udharPersons[i]=u;
    notifyListeners();
  }
  Future<void> deleteUdharPerson(int id) async {
    await _db.deleteUdharPerson(id);
    udharPersons.removeWhere((u)=>u.id==id);
    notifyListeners();
  }
  Future<void> giveUdhar(int uid, double amount, String date, String note) async {
    final u = udharPersons.firstWhere((x)=>x.id==uid, orElse: ()=>UdharPerson(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    u.balance += amount;
    await _db.updateUdharPerson(u);
    final t = AppTransaction(date:date,type:'ادھار دیا',amount:amount,note:note,timestamp:ts);
    await _db.insertUdharTx(uid, t);
    u.transactions.add(t);
    await _addCashLedger(-amount,'ادھار دیا',date,note:note);
    undoStack.add(UndoAction(type:'udhar_give',data:{'uid':uid,'amount':amount,'ts':ts}));
    notifyListeners();
  }
  Future<void> takeUdhar(int uid, double amount, String date, String note) async {
    final u = udharPersons.firstWhere((x)=>x.id==uid, orElse: ()=>UdharPerson(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    u.balance -= amount;
    await _db.updateUdharPerson(u);
    final t = AppTransaction(date:date,type:'ادھار لیا',amount:-amount,note:note,timestamp:ts);
    await _db.insertUdharTx(uid, t);
    u.transactions.add(t);
    await _addCashLedger(amount,'ادھار لیا',date,note:note);
    undoStack.add(UndoAction(type:'udhar_take',data:{'uid':uid,'amount':amount,'ts':ts}));
    notifyListeners();
  }
  Future<void> returnUdhar(int uid, double amount, bool wasGiven, String date, String note) async {
    final u = udharPersons.firstWhere((x)=>x.id==uid, orElse: ()=>UdharPerson(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    final ts = DateTime.now().millisecondsSinceEpoch;
    if (wasGiven) { u.balance -= amount; await _addCashLedger(amount,'ادھار واپسی',date,note:note); }
    else           { u.balance += amount; await _addCashLedger(-amount,'ادھار واپسی',date,note:note); }
    await _db.updateUdharPerson(u);
    final t = AppTransaction(date:date,type:'واپسی',amount:wasGiven?-amount:amount,note:note,timestamp:ts);
    await _db.insertUdharTx(uid, t);
    u.transactions.add(t);
    notifyListeners();
  }

  // ── Stock ───────────────────────────────────────────────────
  Future<void> addStockItem(StockItem item) async {
    final id = await _db.insertStockItem(item);
    stockItems.add(item.copyWith()..id == id);
    final newItem = StockItem(id:id,name:item.name,category:item.category,purchaseRate:item.purchaseRate,quantity:item.quantity,alertLimit:item.alertLimit,lastSupplierId:item.lastSupplierId,purchaseHistory:[],createdAt:item.createdAt,unit:item.unit);
    stockItems.removeWhere((x)=>x.id==0);
    stockItems.add(newItem);
    stockItems.sort((a,b)=>a.name.compareTo(b.name));
    undoStack.add(UndoAction(type:'add_stock',data:{'id':id}));
    notifyListeners();
  }
  Future<void> updateStockItem(StockItem item) async {
    await _db.updateStockItem(item);
    final i = stockItems.indexWhere((x)=>x.id==item.id);
    if (i!=-1) stockItems[i]=item;
    notifyListeners();
  }
  Future<void> deleteStockItem(int id) async {
    await _db.deleteStockItem(id);
    stockItems.removeWhere((i)=>i.id==id);
    notifyListeners();
  }
  Future<void> saveStockAdj({required int itemId, required String adjType,
      required double qty, required double rate, required String date, String note=''}) async {
    final item = stockItems.firstWhere((i)=>i.id==itemId, orElse:()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
    if (adjType=='loss' && qty > item.quantity) {
      throw Exception('وزن کمی $qty ${item.unit} اسٹاک (${item.quantity}) سے زیادہ نہیں ہو سکتی!');
    }
    final amount = qty * rate;
    final ts = DateTime.now().millisecondsSinceEpoch;
    final adj = StockAdjustment(id:0,itemId:itemId,itemName:item.name,adjType:adjType,qty:qty,rate:rate,amount:amount,unit:item.unit,date:date,note:note.isEmpty?(adjType=='loss'?'وزن کمی':'وزن اضافہ'):note,timestamp:ts);
    final id = await _db.insertStockAdjustment(adj);
    stockAdj.insert(0, StockAdjustment(id:id,itemId:itemId,itemName:item.name,adjType:adjType,qty:qty,rate:rate,amount:amount,unit:item.unit,date:date,note:adj.note,timestamp:ts));
    if (adjType=='loss') item.quantity -= qty;
    else item.quantity += qty;
    await _db.updateStockItem(item);
    undoStack.add(UndoAction(type:'stock_adj',data:{'adjId':id,'itemId':itemId,'adjType':adjType,'qty':qty}));
    notifyListeners();
  }

  // ── Sale ────────────────────────────────────────────────────
  Future<int> addSale({required int customerId, required List<SaleItem> items,
      required double discount, required double fee,
      required double cashReceived, required String date, String note=''}) async {
    final ts = DateTime.now().millisecondsSinceEpoch;
    final subTotal = items.fold(0.0,(s,i)=>s+i.total);
    final total    = subTotal - discount + fee;
    final credit   = total - cashReceived;
    final sale = Sale(id:0,customerId:customerId,items:items,total:total,discount:discount,fee:fee,cashReceived:cashReceived,creditAmount:credit,date:date,timestamp:ts,note:note);
    final saleId = await _db.insertSale(sale);
    // stock reduce
    for (final i in items) {
      final item = stockItems.firstWhere((x)=>x.id==i.itemId, orElse:()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
      item.quantity -= i.qty;
      await _db.updateStockItem(item);
    }
    // customer balance
    final c = customers.firstWhere((x)=>x.id==customerId, orElse: ()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    c.balance += credit;
    await _db.updateCustomer(c);
    // customer tx
    final txType = cashReceived<=0?'فروخت (ادھار)':credit<=0?'فروخت (نقد)':'فروخت (جزوی)';
    final ct = AppTransaction(date:date,type:txType,amount:credit,note:note,timestamp:ts,saleId:saleId);
    await _db.insertCustomerTx(customerId, ct);
    c.transactions.add(ct);
    if (cashReceived > 0) {
      await _addCashLedger(cashReceived,'فروخت وصولی',date,note:note);
    }
    final finalSale = Sale(id:saleId,customerId:customerId,items:items,total:total,discount:discount,fee:fee,cashReceived:cashReceived,creditAmount:credit,date:date,timestamp:ts,note:note);
    sales.insert(0, finalSale);
    undoStack.add(UndoAction(type:'sale',data:{'saleId':saleId,'customerId':customerId,'total':total,'cashReceived':cashReceived,'credit':credit,'items':items.map((i)=>i.toMap()).toList()}));
    notifyListeners();
    return saleId;
  }

  Future<void> deleteSale(int saleId) async {
    final sale = sales.firstWhere((s)=>s.id==saleId, orElse: ()=>Sale(id:0,customerId:0,items:[],total:0,discount:0,fee:0,cashReceived:0,creditAmount:0,date:'',timestamp:0));
    for (final i in sale.items) {
      final item = stockItems.firstWhere((x)=>x.id==i.itemId, orElse:()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
      if (item.id != 0) { item.quantity += i.qty; await _db.updateStockItem(item); }
    }
    final c = customers.firstWhere((x)=>x.id==sale.customerId, orElse:()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    if (c.id != 0) {
      c.balance -= sale.creditAmount;
      await _db.updateCustomer(c);
      await _db.deleteCustomerTxBySaleId(c.id, saleId);
      c.transactions.removeWhere((t)=>t.saleId==saleId);
    }
    if (sale.cashReceived > 0) await _addCashLedger(-sale.cashReceived,'فروخت حذف',sale.date);
    await _db.deleteSale(saleId);
    sales.removeWhere((s)=>s.id==saleId);
    notifyListeners();
  }

  // ── Purchase ────────────────────────────────────────────────
  Future<int> addPurchase({required int supplierId, required List<PurchaseItem> items,
      required double discount, required double fee,
      required double cashPaid, required String date, String note=''}) async {
    final ts = DateTime.now().millisecondsSinceEpoch;
    final subTotal = items.fold(0.0,(s,i)=>s+i.total);
    final total    = subTotal - discount + fee;
    final credit   = total - cashPaid;
    final p = Purchase(id:0,supplierId:supplierId,items:items,total:total,discount:discount,fee:fee,cashPaid:cashPaid,creditAmount:credit,date:date,timestamp:ts,note:note);
    final pId = await _db.insertPurchase(p);
    // stock add + WAC
    for (final i in items) {
      final item = stockItems.firstWhere((x)=>x.id==i.itemId, orElse:()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
      final newWac = _newWac(i.itemId, i.qty, i.rate);
      item.quantity     += i.qty;
      item.purchaseRate  = newWac;
      item.lastSupplierId = supplierId;
      await _db.updateStockItem(item);
      final ph = PurchaseHistory(supplierId:supplierId,rate:i.rate,quantity:i.qty,date:date,purchaseId:pId,timestamp:ts);
      await _db.insertPurchaseHistory(i.itemId, ph);
      item.purchaseHistory.add(ph);
    }
    // supplier balance
    final s = suppliers.firstWhere((x)=>x.id==supplierId, orElse: ()=>Supplier(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    s.balance += credit;
    await _db.updateSupplier(s);
    final txType = cashPaid<=0?'خریداری (بقایا)':credit<=0?'خریداری (نقد)':'خریداری (جزوی)';
    final st = AppTransaction(date:date,type:txType,amount:credit,note:note,timestamp:ts,purchaseId:pId);
    await _db.insertSupplierTx(supplierId, st);
    s.transactions.add(st);
    if (cashPaid > 0) await _addCashLedger(-cashPaid,'خریداری ادائیگی',date,note:note);
    purchases.insert(0, Purchase(id:pId,supplierId:supplierId,items:items,total:total,discount:discount,fee:fee,cashPaid:cashPaid,creditAmount:credit,date:date,timestamp:ts,note:note));
    undoStack.add(UndoAction(type:'purchase',data:{'pId':pId,'supplierId':supplierId,'total':total,'cashPaid':cashPaid,'credit':credit}));
    notifyListeners();
    return pId;
  }

  Future<void> deletePurchase(int pId) async {
    final p = purchases.firstWhere((x)=>x.id==pId, orElse: ()=>Purchase(id:0,supplierId:0,items:[],total:0,discount:0,fee:0,cashPaid:0,creditAmount:0,date:'',timestamp:0));
    for (final i in p.items) {
      final item = stockItems.firstWhere((x)=>x.id==i.itemId, orElse:()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
      if (item.id != 0) {
        item.quantity -= i.qty;
        await _db.updateStockItem(item);
        await _db.deleteLastPurchaseHistory(i.itemId);
        if (item.purchaseHistory.isNotEmpty) item.purchaseHistory.removeLast();
      }
    }
    final s = suppliers.firstWhere((x)=>x.id==p.supplierId, orElse:()=>Supplier(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
    if (s.id != 0) {
      s.balance -= p.creditAmount;
      await _db.updateSupplier(s);
      await _db.deleteSupplierTxByPurchaseId(s.id, pId);
      s.transactions.removeWhere((t)=>t.purchaseId==pId);
    }
    if (p.cashPaid > 0) await _addCashLedger(p.cashPaid,'خریداری حذف',p.date);
    await _db.deletePurchase(pId);
    purchases.removeWhere((x)=>x.id==pId);
    notifyListeners();
  }

  // ── Expenses ────────────────────────────────────────────────
  Future<void> addExpense({required String category, required double amount,
      required String date, String note='', int? linkedSupplierId}) async {
    final ts = DateTime.now().millisecondsSinceEpoch;
    final e = Expense(id:0,category:category,amount:amount,note:note,date:date,timestamp:ts,linkedSupplierId:linkedSupplierId);
    final id = await _db.insertExpense(e);
    expenses.insert(0, Expense(id:id,category:category,amount:amount,note:note,date:date,timestamp:ts,linkedSupplierId:linkedSupplierId));
    await _addCashLedger(-amount,'خرچہ: $category',date,note:note);
    undoStack.add(UndoAction(type:'expense',data:{'id':id,'amount':amount}));
    notifyListeners();
  }
  Future<void> deleteExpense(int id) async {
    final e = expenses.firstWhere((x)=>x.id==id, orElse: ()=>Expense(id:0,category:'',amount:0,date:'',timestamp:0));
    await _db.deleteExpense(id);
    await _addCashLedger(e.amount,'خرچہ حذف',e.date);
    expenses.removeWhere((x)=>x.id==id);
    notifyListeners();
  }

  // ── Cash In/Out ─────────────────────────────────────────────
  Future<void> addCashIn(double amount, String date, String note) async {
    final ts = DateTime.now().millisecondsSinceEpoch;
    final entry = CashInOutEntry(id:0,type:'کیش ان',amount:amount,note:note,date:date,timestamp:ts);
    final id = await _db.insertCashInOut(entry);
    cashInOutList.insert(0, CashInOutEntry(id:id,type:'کیش ان',amount:amount,note:note,date:date,timestamp:ts));
    await _addCashLedger(amount,'کیش ان',date,note:note);
    undoStack.add(UndoAction(type:'cash_in',data:{'id':id,'amount':amount}));
    notifyListeners();
  }
  Future<void> addCashOut(double amount, String date, String note) async {
    final ts = DateTime.now().millisecondsSinceEpoch;
    final entry = CashInOutEntry(id:0,type:'کیش آؤٹ',amount:-amount,note:note,date:date,timestamp:ts);
    final id = await _db.insertCashInOut(entry);
    cashInOutList.insert(0, CashInOutEntry(id:id,type:'کیش آؤٹ',amount:-amount,note:note,date:date,timestamp:ts));
    await _addCashLedger(-amount,'کیش آؤٹ',date,note:note);
    undoStack.add(UndoAction(type:'cash_out',data:{'id':id,'amount':amount}));
    notifyListeners();
  }
  Future<void> deleteCashEntry(int id) async {
    final e = cashInOutList.firstWhere((x)=>x.id==id, orElse: ()=>CashInOutEntry(id:0,type:'',amount:0,date:'',timestamp:0));
    await _db.deleteCashInOut(id);
    await _addCashLedger(-e.amount,'کیش انٹری حذف',e.date);
    cashInOutList.removeWhere((x)=>x.id==id);
    notifyListeners();
  }

  // ── Partners ────────────────────────────────────────────────
  Future<void> addPartner(Partner p) async {
    final id = await _db.insertPartner(p);
    partners.add(Partner(id:id,name:p.name,sharePercent:p.sharePercent,totalWithdrawal:0,withdrawals:[]));
    notifyListeners();
  }
  Future<void> updatePartner(Partner p) async {
    await _db.updatePartner(p);
    final i = partners.indexWhere((x)=>x.id==p.id);
    if (i!=-1) partners[i]=p;
    notifyListeners();
  }
  Future<void> deletePartner(int id) async {
    await _db.deletePartner(id);
    partners.removeWhere((p)=>p.id==id);
    notifyListeners();
  }
  Future<void> partnerWithdrawal(int pid, double amount, String note, String date) async {
    final p = partners.firstWhere((x)=>x.id==pid, orElse: ()=>Partner(id:0,name:'',sharePercent:0,totalWithdrawal:0,withdrawals:[]));
    final ts = DateTime.now().millisecondsSinceEpoch;
    p.totalWithdrawal += amount;
    await _db.updatePartner(p);
    await _db.insertPartnerWithdrawal(pid, amount, note, date, ts);
    p.withdrawals.insert(0, {'amount':amount,'note':note,'date':date,'timestamp':ts});
    await _addCashLedger(-amount,'پارٹنر نکاسی: ${p.name}',date);
    notifyListeners();
  }
  double getPartnerShare(int pid) {
    try {
      final p = partners.firstWhere((x)=>x.id==pid, orElse: ()=>Partner(id:0,name:'',sharePercent:0,totalWithdrawal:0,withdrawals:[]));
      return netProfit * (p.sharePercent / 100);
    } catch (_) { return 0; }
  }

  // ── Undo ────────────────────────────────────────────────────
  Future<bool> undo() async {
    if (undoStack.isEmpty) return false;
    final action = undoStack.removeLast();
    try {
      switch (action.type) {
        case 'sale':     await deleteSale(action.data['saleId'] as int); undoStack.removeLast(); break;
        case 'purchase': await deletePurchase(action.data['pId'] as int); undoStack.removeLast(); break;
        case 'expense':  await deleteExpense(action.data['id'] as int); undoStack.removeLast(); break;
        case 'cash_in':  await deleteCashEntry(action.data['id'] as int); undoStack.removeLast(); break;
        case 'cash_out': await deleteCashEntry(action.data['id'] as int); undoStack.removeLast(); break;
        case 'cust_receipt':
          final c = customers.firstWhere((x)=>x.id==action.data['cid'], orElse:()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          c.balance += (action.data['amount'] as num).toDouble();
          await _db.updateCustomer(c);
          await _db.deleteLastCustomerTx(c.id);
          c.transactions.removeLast();
          await _addCashLedger(-(action.data['amount'] as num).toDouble(),'undo وصولی',_today());
          break;
        case 'cust_disc':
          final c = customers.firstWhere((x)=>x.id==action.data['cid'], orElse:()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          c.balance += (action.data['amount'] as num).toDouble();
          await _db.updateCustomer(c);
          await _db.deleteLastCustomerTx(c.id);
          c.transactions.removeLast();
          break;
        case 'cust_tax':
          final c = customers.firstWhere((x)=>x.id==action.data['cid'], orElse:()=>Customer(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          c.balance -= (action.data['amount'] as num).toDouble();
          await _db.updateCustomer(c);
          await _db.deleteLastCustomerTx(c.id);
          c.transactions.removeLast();
          break;
        case 'supp_pay':
          final s = suppliers.firstWhere((x)=>x.id==action.data['sid'], orElse:()=>Supplier(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          s.balance += (action.data['amount'] as num).toDouble();
          await _db.updateSupplier(s);
          await _db.deleteLastSupplierTx(s.id);
          s.transactions.removeLast();
          await _addCashLedger((action.data['amount'] as num).toDouble(),'undo ادائیگی',_today());
          break;
        case 'supp_disc':
          final s = suppliers.firstWhere((x)=>x.id==action.data['sid'], orElse:()=>Supplier(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          s.balance += (action.data['amount'] as num).toDouble();
          await _db.updateSupplier(s);
          await _db.deleteLastSupplierTx(s.id);
          s.transactions.removeLast();
          break;
        case 'udhar_give':
          final u = udharPersons.firstWhere((x)=>x.id==action.data['uid'], orElse:()=>UdharPerson(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          u.balance -= (action.data['amount'] as num).toDouble();
          await _db.updateUdharPerson(u);
          await _db.deleteLastUdharTx(u.id);
          u.transactions.removeLast();
          await _addCashLedger((action.data['amount'] as num).toDouble(),'undo ادھار',_today());
          break;
        case 'udhar_take':
          final u = udharPersons.firstWhere((x)=>x.id==action.data['uid'], orElse:()=>UdharPerson(id:0,name:'',phone:'',address:'',balance:0,transactions:[],createdAt:''));
          u.balance += (action.data['amount'] as num).toDouble();
          await _db.updateUdharPerson(u);
          await _db.deleteLastUdharTx(u.id);
          u.transactions.removeLast();
          await _addCashLedger(-(action.data['amount'] as num).toDouble(),'undo ادھار',_today());
          break;
        case 'add_customer': await deleteCustomer(action.data['id'] as int); undoStack.removeLast(); break;
        case 'add_supplier': await deleteSupplier(action.data['id'] as int); undoStack.removeLast(); break;
        case 'add_stock':    await deleteStockItem(action.data['id'] as int); undoStack.removeLast(); break;
        case 'stock_adj':
          await _db.deleteStockAdjustment(action.data['adjId'] as int);
          final item = stockItems.firstWhere((i)=>i.id==action.data['itemId'], orElse:()=>StockItem(id:0,name:'',category:'',purchaseRate:0,quantity:0,alertLimit:0,purchaseHistory:[],createdAt:''));
          final qty = (action.data['qty'] as num).toDouble();
          if (action.data['adjType']=='loss') item.quantity += qty;
          else item.quantity -= qty;
          await _db.updateStockItem(item);
          stockAdj.removeWhere((a)=>a.id==action.data['adjId']);
          break;
      }
      notifyListeners();
      return true;
    } catch (_) { return false; }
  }

  // ── Backup ──────────────────────────────────────────────────
  Future<String> exportBackup() async {
    final data = await _db.exportAll();
    final json = jsonEncode(data);
    final dir  = await getApplicationDocumentsDirectory();
    final now  = DateTime.now();
    final fn   = 'iqbal_backup_${now.year}${now.month.toString().padLeft(2,'0')}${now.day.toString().padLeft(2,'0')}.json';
    final file = File('${dir.path}/$fn');
    await file.writeAsString(json);
    return file.path;
  }
  Future<bool> importBackup() async {
    final res = await FilePicker.platform.pickFiles(type:FileType.custom,allowedExtensions:['json']);
    if (res==null||res.files.isEmpty) return false;
    final path = res.files.first.path;
    if (path==null) return false;
    final json = await File(path).readAsString();
    await _db.importAll(jsonDecode(json) as Map<String, dynamic>);
    await _reload();
    notifyListeners();
    return true;
  }
  Future<void> clearTransactions() async {
    await _db.clearTransactions(); await _reload(); notifyListeners();
  }
  Future<void> fullReset() async {
    await _db.fullReset(); await _reload(); notifyListeners();
  }

  // ── Filters ─────────────────────────────────────────────────
  List<Sale>     getSalesByRange(String from, String to) => sales.where((s)=>s.date>=from&&s.date<=to).toList();
  List<Purchase> getPurchByRange(String from, String to) => purchases.where((p)=>p.date>=from&&p.date<=to).toList();
  List<Expense>  getExpByRange(String from, String to)   => expenses.where((e)=>e.date>=from&&e.date<=to).toList();

  // ── Lookups ─────────────────────────────────────────────────
  StockItem?   getStock(int id) => stockItems.cast<StockItem?>().firstWhere((i)=>i?.id==id, orElse:()=>null);
  Customer?    getCust(int id) => customers.cast<Customer?>().firstWhere((c)=>c?.id==id, orElse:()=>null);
  Supplier?    getSupp(int id) => suppliers.cast<Supplier?>().firstWhere((s)=>s?.id==id, orElse:()=>null);

  // aliases for reports screen compatibility
  List<Sale>     getSalesByDateRange(String from, String to) => getSalesByRange(from, to);
  List<Purchase> getPurchasesByDateRange(String from, String to) => getPurchByRange(from, to);
  List<Expense>  getExpensesByDateRange(String from, String to) => getExpByRange(from, to);

  // alias for dashboard screen
  List<Expense> get todayExpTotalList => todayExpList;
}


// ============================================================
// THEME
// ============================================================

void configureGoogleFonts() {
  GoogleFonts.config.allowRuntimeFetching = false;
}

class AppTheme {
  static const Color primary     = Color(0xFF667eea);
  static const Color primaryDark = Color(0xFF764ba2);
  static const Color success     = Color(0xFF28a745);
  static const Color danger      = Color(0xFFdc3545);
  static const Color warning     = Color(0xFFffc107);
  static const Color info        = Color(0xFF17a2b8);
  static const Color orange      = Color(0xFFfd7e14);
  static const Color purple      = Color(0xFF6f42c1);
  static const Color lightBg     = Color(0xFFF8F9FA);
  static const Color cardBorder  = Color(0xFFDEE2E6);
  static const Color textPrimary = Color(0xFF495057);

  static TextStyle urduStyle({
    double fontSize = 14,
    FontWeight weight = FontWeight.normal,
    Color color = textPrimary,
  }) {
    try {
      return GoogleFonts.notoNastaliqUrdu(fontSize: fontSize, fontWeight: weight, color: color);
    } catch (_) {
      return TextStyle(fontSize: fontSize, fontWeight: weight, color: color);
    }
  }

  static ThemeData get theme {
    final base = ThemeData(
      useMaterial3: true,
      colorScheme: ColorScheme.fromSeed(
        seedColor: primary, primary: primary,
        secondary: primaryDark, error: danger, surface: Colors.white,
      ),
    );
    TextTheme urduTextTheme;
    try {
      urduTextTheme = GoogleFonts.notoNastaliqUrduTextTheme(base.textTheme).copyWith(
        bodyMedium: GoogleFonts.notoNastaliqUrdu(color: textPrimary, fontSize: 14),
        bodySmall:  GoogleFonts.notoNastaliqUrdu(color: textPrimary, fontSize: 12),
      );
    } catch (_) { urduTextTheme = base.textTheme; }

    return base.copyWith(
      textTheme: urduTextTheme,
      appBarTheme: const AppBarTheme(
        backgroundColor: primary, foregroundColor: Colors.white,
        elevation: 0, centerTitle: true,
        titleTextStyle: TextStyle(fontSize: 20, fontWeight: FontWeight.bold, color: Colors.white),
      ),
      cardTheme: CardThemeData(
        elevation: 2, margin: EdgeInsets.zero,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      ),
      elevatedButtonTheme: ElevatedButtonThemeData(
        style: ElevatedButton.styleFrom(
          backgroundColor: primary, foregroundColor: Colors.white,
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
          padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 12),
          textStyle: const TextStyle(fontSize: 15, fontWeight: FontWeight.bold),
        ),
      ),
      outlinedButtonTheme: OutlinedButtonThemeData(
        style: OutlinedButton.styleFrom(
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
          padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
        ),
      ),
      inputDecorationTheme: InputDecorationTheme(
        border: OutlineInputBorder(
          borderRadius: BorderRadius.circular(8),
          borderSide: const BorderSide(color: cardBorder),
        ),
        focusedBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(8),
          borderSide: const BorderSide(color: primary, width: 2),
        ),
        contentPadding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
        labelStyle: const TextStyle(color: textPrimary, fontWeight: FontWeight.bold),
        hintStyle: TextStyle(color: Colors.grey.shade400),
      ),
      scaffoldBackgroundColor: lightBg,
      dividerTheme: const DividerThemeData(color: cardBorder, thickness: 0.8),
      snackBarTheme: SnackBarThemeData(
        behavior: SnackBarBehavior.floating,
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
      ),
    );
  }
}


// ============================================================
// IQBAL TRADERS — Excel Export Service
// تمام رپورٹس Excel میں برآمد
// ============================================================

class ExcelExportService {
  // ============================================================
  // فروخت رپورٹ
  // ============================================================
  static Future<String> exportSalesReport({
    required List<Sale> sales,
    required List<Customer> customers,
    required List<StockItem> stockItems,
    required String shopName,
    required String dateRange,
  }) async {
    final excel = Excel.createExcel();
    final sheet = excel['فروخت رپورٹ'];

    // ہیڈر اسٹائل
    final headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.fromHexString('#667eea'),
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    );

    // ہیڈر رو
    final headers = ['رسید #', 'تاریخ', 'کسٹمر', 'آئٹمز', 'سب ٹوٹل', 'ڈسکاؤنٹ', 'اضافی', 'کل', 'نقد', 'بقایا', 'نوٹ'];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      cell.value = TextCellValue(headers[i]);
      cell.cellStyle = headerStyle;
    }

    // ڈیٹا رو
    double totalAmount = 0, totalCash = 0, totalCredit = 0;
    for (int r = 0; r < sales.length; r++) {
      final sale = sales[r];
      final customer = customers.firstWhere(
        (c) => c.id == sale.customerId,
        orElse: () => Customer(id: 0, name: '-', phone: '', address: '', balance: 0, transactions: [], createdAt: ''),
      );
      final itemDesc = sale.items.map((si) {
        final item = stockItems.firstWhere((i) => i.id == si.itemId, orElse: () => StockItem(id: 0, name: '?', category: '', purchaseRate: 0, quantity: 0, alertLimit: 0, purchaseHistory: [], createdAt: ''),
        );
        return '${item.name} ×${si.qty}${item.unit}@${si.rate.toStringAsFixed(2)}';
      }).join(' | ');

      final subTotal = sale.items.fold(0.0, (s, i) => s + i.total);
      final rowData = [
        sale.id.toString(),
        sale.date,
        customer.name,
        itemDesc,
        subTotal.toStringAsFixed(2),
        sale.discount.toStringAsFixed(2),
        sale.fee.toStringAsFixed(2),
        sale.total.toStringAsFixed(2),
        sale.cashReceived.toStringAsFixed(2),
        sale.creditAmount.toStringAsFixed(2),
        sale.note ?? '',
      ];

      totalAmount += sale.total;
      totalCash += sale.cashReceived;
      totalCredit += sale.creditAmount;

      for (int c = 0; c < rowData.length; c++) {
        final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1));
        cell.value = TextCellValue(rowData[c]);
        if (r % 2 == 0) {
          cell.cellStyle = CellStyle(backgroundColorHex: ExcelColor.fromHexString('#F8F9FA'));
        }
      }
    }

    // ٹوٹل رو
    final totalRow = sales.length + 1;
    final totalStyle = CellStyle(bold: true, backgroundColorHex: ExcelColor.fromHexString('#E8F5E9'));
    final totals = ['', '', 'مجموعہ', '', '', '', '', totalAmount.toStringAsFixed(2), totalCash.toStringAsFixed(2), totalCredit.toStringAsFixed(2), ''];
    for (int c = 0; c < totals.length; c++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: totalRow));
      cell.value = TextCellValue(totals[c]);
      cell.cellStyle = totalStyle;
    }

    // فائل محفوظ
    return await _saveExcel(excel, 'Sales_Report_${DateTime.now().millisecondsSinceEpoch}');
  }

  // ============================================================
  // خریداری رپورٹ
  // ============================================================
  static Future<String> exportPurchaseReport({
    required List<Purchase> purchases,
    required List<Supplier> suppliers,
    required List<StockItem> stockItems,
  }) async {
    final excel = Excel.createExcel();
    final sheet = excel['خریداری رپورٹ'];

    final headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.fromHexString('#fd7e14'),
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    );

    final headers = ['#', 'تاریخ', 'سپلائر', 'آئٹمز', 'کل', 'ڈسکاؤنٹ', 'ادا', 'بقایا', 'نوٹ'];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      cell.value = TextCellValue(headers[i]);
      cell.cellStyle = headerStyle;
    }

    for (int r = 0; r < purchases.length; r++) {
      final p = purchases[r];
      final supplier = suppliers.firstWhere(
        (s) => s.id == p.supplierId,
        orElse: () => Supplier(id: 0, name: '-', phone: '', address: '', balance: 0, transactions: [], createdAt: ''),
      );
      final itemDesc = p.items.map((pi) {
        final item = stockItems.firstWhere(
          (i) => i.id == pi.itemId,
          orElse: () => StockItem(id: 0, name: '?', category: '', purchaseRate: 0, quantity: 0, alertLimit: 0, purchaseHistory: [], createdAt: ''),
        );
        return '${item.name} ×${pi.qty}${item.unit}@${pi.rate.toStringAsFixed(2)}';
      }).join(' | ');

      final rowData = [
        p.id.toString(), p.date, supplier.name, itemDesc,
        p.total.toStringAsFixed(2), p.discount.toStringAsFixed(2),
        p.cashPaid.toStringAsFixed(2), p.creditAmount.toStringAsFixed(2), p.note ?? '',
      ];

      for (int c = 0; c < rowData.length; c++) {
        final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1));
        cell.value = TextCellValue(rowData[c]);
      }
    }

    return await _saveExcel(excel, 'Purchases_Report_${DateTime.now().millisecondsSinceEpoch}');
  }

  // ============================================================
  // اخراجات رپورٹ
  // ============================================================
  static Future<String> exportExpensesReport({
    required List<Expense> expenses,
  }) async {
    final excel = Excel.createExcel();
    final sheet = excel['اخراجات'];

    final headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.fromHexString('#dc3545'),
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    );

    final headers = ['#', 'تاریخ', 'کیٹیگری', 'رقم', 'نوٹ'];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      cell.value = TextCellValue(headers[i]);
      cell.cellStyle = headerStyle;
    }

    // کیٹیگری وار خلاصہ شیٹ
    final summarySheet = excel['کیٹیگری خلاصہ'];
    final Map<String, double> catTotals = {};
    for (final e in expenses) {
      catTotals[e.category] = (catTotals[e.category] ?? 0) + e.amount;
    }

    int sr = 0;
    for (final entry in catTotals.entries.toList()..sort((a, b) => b.value.compareTo(a.value))) {
      summarySheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: sr)).value = TextCellValue(entry.key);
      summarySheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: sr)).value = TextCellValue(entry.value.toStringAsFixed(2));
      sr++;
    }

    double total = 0;
    for (int r = 0; r < expenses.length; r++) {
      final e = expenses[r];
      total += e.amount;
      final rowData = [(r + 1).toString(), e.date, e.category, e.amount.toStringAsFixed(2), e.note ?? ''];
      for (int c = 0; c < rowData.length; c++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1)).value = TextCellValue(rowData[c]);
      }
    }

    // ٹوٹل
    final t = expenses.length + 1;
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: t)).value = TextCellValue('مجموعہ');
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: t)).value = TextCellValue(total.toStringAsFixed(2));

    return await _saveExcel(excel, 'Expenses_Report_${DateTime.now().millisecondsSinceEpoch}');
  }

  // ============================================================
  // اسٹاک رپورٹ
  // ============================================================
  static Future<String> exportStockReport({
    required List<StockItem> stockItems,
  }) async {
    final excel = Excel.createExcel();
    final sheet = excel['اسٹاک'];

    final headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.fromHexString('#17a2b8'),
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    );

    final headers = ['#', 'آئٹم', 'کیٹیگری', 'مقدار', 'یونٹ', 'WAC ریٹ', 'مالیت', 'الرٹ حد', 'حالت'];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      cell.value = TextCellValue(headers[i]);
      cell.cellStyle = headerStyle;
    }

    double totalValue = 0;
    for (int r = 0; r < stockItems.length; r++) {
      final item = stockItems[r];
      final value = item.quantity * item.purchaseRate;
      totalValue += value;

      final rowData = [
        (r + 1).toString(), item.name, item.category,
        item.quantity.toStringAsFixed(2), item.unit,
        item.purchaseRate.toStringAsFixed(2), value.toStringAsFixed(2),
        item.alertLimit.toStringAsFixed(2),
        item.isLowStock ? 'کم اسٹاک ⚠️' : 'ٹھیک ✅',
      ];

      for (int c = 0; c < rowData.length; c++) {
        final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r + 1));
        cell.value = TextCellValue(rowData[c]);
        if (item.isLowStock) {
          cell.cellStyle = CellStyle(backgroundColorHex: ExcelColor.fromHexString('#FFF3CD'));
        }
      }
    }

    // ٹوٹل
    final t = stockItems.length + 1;
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: t)).value = TextCellValue('مجموعہ مالیت');
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 6, rowIndex: t)).value = TextCellValue(totalValue.toStringAsFixed(2));

    return await _saveExcel(excel, 'Stock_Report_${DateTime.now().millisecondsSinceEpoch}');
  }

  // ============================================================
  // کسٹمرز رپورٹ
  // ============================================================
  static Future<String> exportCustomersReport({
    required List<Customer> customers,
  }) async {
    final excel = Excel.createExcel();
    final sheet = excel['کسٹمرز'];

    final headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.fromHexString('#667eea'),
      fontColorHex: ExcelColor.fromHexString('#FFFFFF'),
    );

    final headers = ['#', 'نام', 'فون', 'پتہ', 'بیلنس', 'حالت'];
    for (int i = 0; i < headers.length; i++) {
      final cell = sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0));
      cell.value = TextCellValue(headers[i]);
      cell.cellStyle = headerStyle;
    }

    for (int r = 0; r < customers.length; r++) {
      final c = customers[r];
      final rowData = [
        (r + 1).toString(), c.name, c.phone, c.address,
        c.balance.toStringAsFixed(2),
        c.balance > 0 ? 'واجب الادا' : c.balance < 0 ? 'اضافی' : 'برابر',
      ];
      for (int col = 0; col < rowData.length; col++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: r + 1)).value = TextCellValue(rowData[col]);
      }
    }

    return await _saveExcel(excel, 'Customers_Report_${DateTime.now().millisecondsSinceEpoch}');
  }

  // ============================================================
  // مکمل رپورٹ (تمام شیٹس)
  // ============================================================
  static Future<String> exportCompleteReport({
    required List<Sale> sales,
    required List<Purchase> purchases,
    required List<Expense> expenses,
    required List<Customer> customers,
    required List<Supplier> suppliers,
    required List<StockItem> stockItems,
    required double cashInHand,
    required double netProfit,
    required String shopName,
  }) async {
    final excel = Excel.createExcel();

    // خلاصہ شیٹ
    final summary = excel['خلاصہ'];
    final headerStyle = CellStyle(bold: true, backgroundColorHex: ExcelColor.fromHexString('#343a40'), fontColorHex: ExcelColor.fromHexString('#FFFFFF'));

    summary.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value = TextCellValue(shopName);
    summary.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).cellStyle = CellStyle(bold: true, fontSize: 14);
    summary.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1)).value = TextCellValue('رپورٹ تاریخ: ${DateTime.now().toIso8601String().split('T')[0]}');

    final summaryData = [
      ['کل فروخت', sales.fold(0.0, (s, x) => s + x.total).toStringAsFixed(2)],
      ['کل خریداری', purchases.fold(0.0, (s, x) => s + x.total).toStringAsFixed(2)],
      ['کل اخراجات', expenses.fold(0.0, (s, x) => s + x.amount).toStringAsFixed(2)],
      ['کیش ان ہینڈ', cashInHand.toStringAsFixed(2)],
      ['خالص منافع', netProfit.toStringAsFixed(2)],
      ['کسٹمر واجبات', customers.fold(0.0, (s, c) => s + (c.balance > 0 ? c.balance : 0)).toStringAsFixed(2)],
      ['سپلائر واجبات', suppliers.fold(0.0, (s, x) => s + (x.balance > 0 ? x.balance : 0)).toStringAsFixed(2)],
      ['اسٹاک مالیت', stockItems.fold(0.0, (s, i) => s + (i.quantity * i.purchaseRate)).toStringAsFixed(2)],
    ];

    for (int r = 0; r < summaryData.length; r++) {
      summary.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r + 3)).value = TextCellValue(summaryData[r][0]);
      summary.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: r + 3)).value = TextCellValue('Rs. ${summaryData[r][1]}');
    }

    // فروخت شیٹ
    _addSalesSheet(excel, sales, customers, stockItems);

    // خریداری شیٹ
    _addPurchasesSheet(excel, purchases, suppliers, stockItems);

    // اخراجات شیٹ
    _addExpensesSheet(excel, expenses);

    // اسٹاک شیٹ
    _addStockSheet(excel, stockItems);

    return await _saveExcel(excel, 'Complete_Report_${DateTime.now().millisecondsSinceEpoch}');
  }

  static void _addSalesSheet(Excel excel, List<Sale> sales, List<Customer> customers, List<StockItem> stockItems) {
    final sheet = excel['فروخت'];
    final headers = ['تاریخ', 'کسٹمر', 'کل', 'نقد', 'بقایا'];
    for (int i = 0; i < headers.length; i++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0)).value = TextCellValue(headers[i]);
    }
    for (int r = 0; r < sales.length; r++) {
      final s = sales[r];
      final c = customers.firstWhere((x) => x.id == s.customerId, orElse: () => Customer(id: 0, name: '-', phone: '', address: '', balance: 0, transactions: [], createdAt: ''));
      final row = [s.date, c.name, s.total.toStringAsFixed(2), s.cashReceived.toStringAsFixed(2), s.creditAmount.toStringAsFixed(2)];
      for (int col = 0; col < row.length; col++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: r + 1)).value = TextCellValue(row[col]);
      }
    }
  }

  static void _addPurchasesSheet(Excel excel, List<Purchase> purchases, List<Supplier> suppliers, List<StockItem> stockItems) {
    final sheet = excel['خریداری'];
    final headers = ['تاریخ', 'سپلائر', 'کل', 'ادا', 'بقایا'];
    for (int i = 0; i < headers.length; i++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0)).value = TextCellValue(headers[i]);
    }
    for (int r = 0; r < purchases.length; r++) {
      final p = purchases[r];
      final s = suppliers.firstWhere((x) => x.id == p.supplierId, orElse: () => Supplier(id: 0, name: '-', phone: '', address: '', balance: 0, transactions: [], createdAt: ''));
      final row = [p.date, s.name, p.total.toStringAsFixed(2), p.cashPaid.toStringAsFixed(2), p.creditAmount.toStringAsFixed(2)];
      for (int col = 0; col < row.length; col++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: r + 1)).value = TextCellValue(row[col]);
      }
    }
  }

  static void _addExpensesSheet(Excel excel, List<Expense> expenses) {
    final sheet = excel['اخراجات'];
    final headers = ['تاریخ', 'کیٹیگری', 'رقم', 'نوٹ'];
    for (int i = 0; i < headers.length; i++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0)).value = TextCellValue(headers[i]);
    }
    for (int r = 0; r < expenses.length; r++) {
      final e = expenses[r];
      final row = [e.date, e.category, e.amount.toStringAsFixed(2), e.note ?? ''];
      for (int col = 0; col < row.length; col++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: r + 1)).value = TextCellValue(row[col]);
      }
    }
  }

  static void _addStockSheet(Excel excel, List<StockItem> stockItems) {
    final sheet = excel['اسٹاک'];
    final headers = ['آئٹم', 'کیٹیگری', 'مقدار', 'یونٹ', 'WAC ریٹ', 'مالیت'];
    for (int i = 0; i < headers.length; i++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: i, rowIndex: 0)).value = TextCellValue(headers[i]);
    }
    for (int r = 0; r < stockItems.length; r++) {
      final i = stockItems[r];
      final row = [i.name, i.category, i.quantity.toStringAsFixed(2), i.unit, i.purchaseRate.toStringAsFixed(2), (i.quantity * i.purchaseRate).toStringAsFixed(2)];
      for (int col = 0; col < row.length; col++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: r + 1)).value = TextCellValue(row[col]);
      }
    }
  }

  // ============================================================
  // فائل محفوظ اور شیئر
  // ============================================================
  static Future<String> _saveExcel(Excel excel, String filename) async {
    final dir = await getApplicationDocumentsDirectory();
    final path = '${dir.path}/$filename.xlsx';
    final file = File(path);
    final bytes = excel.save();
    if (bytes != null) {
      await file.writeAsBytes(bytes);
    }
    return path;
  }

  static Future<void> shareFile(String path, {String? message}) async {
    await Share.shareXFiles(
      [XFile(path)],
      text: message ?? 'IQBAL TRADERS رپورٹ',
    );
  }
}


// ============================================================
// IQBAL TRADERS — پرنٹنگ سروس
// تھرمل رسید + A4 رپورٹ
// ============================================================

class PrintService {
  // ============================================================
  // فروخت رسید — تھرمل (58mm / 80mm)
  // ============================================================
  static Future<void> printSaleReceipt({
    required Sale sale,
    required Customer customer,
    required List<StockItem> stockItems,
    required String shopName,
  }) async {
    final doc = pw.Document();

    doc.addPage(pw.Page(
      pageFormat: PdfPageFormat(
        58 * PdfPageFormat.mm,
        double.infinity,
        marginAll: 4 * PdfPageFormat.mm,
      ),
      build: (pw.Context ctx) {
        return pw.Column(
          crossAxisAlignment: pw.CrossAxisAlignment.center,
          children: [
            // ہیڈر
            pw.Text(shopName,
                style: pw.TextStyle(fontSize: 14, fontWeight: pw.FontWeight.bold)),
            pw.SizedBox(height: 2),
            pw.Text('فروخت رسید',
                style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold)),
            pw.Divider(thickness: 1),

            // تاریخ اور نمبر
            pw.Row(
              mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
              children: [
                pw.Text('تاریخ: ${sale.date}', style: const pw.TextStyle(fontSize: 8)),
                pw.Text('رسید #${sale.id}', style: const pw.TextStyle(fontSize: 8)),
              ],
            ),
            pw.Text('کسٹمر: ${customer.name}',
                style: const pw.TextStyle(fontSize: 9)),
            pw.Divider(thickness: 0.5),

            // آئٹمز
            pw.Table(
              columnWidths: {
                0: const pw.FlexColumnWidth(3),
                1: const pw.FlexColumnWidth(1.5),
                2: const pw.FlexColumnWidth(1.5),
                3: const pw.FlexColumnWidth(2),
              },
              children: [
                // ہیڈر رو
                pw.TableRow(children: [
                  pw.Text('آئٹم', style: pw.TextStyle(fontSize: 8, fontWeight: pw.FontWeight.bold)),
                  pw.Text('مقدار', style: pw.TextStyle(fontSize: 8, fontWeight: pw.FontWeight.bold)),
                  pw.Text('ریٹ', style: pw.TextStyle(fontSize: 8, fontWeight: pw.FontWeight.bold)),
                  pw.Text('کل', style: pw.TextStyle(fontSize: 8, fontWeight: pw.FontWeight.bold), textAlign: pw.TextAlign.right),
                ]),
                // آئٹم رو
                ...sale.items.map((item) {
                  final stockItem = stockItems.firstWhere((i, orElse: () => _empty()) => i.id == item.itemId, orElse: () => StockItem(id: 0, name: '?', category: '', purchaseRate: 0, quantity: 0, alertLimit: 0, purchaseHistory: [], createdAt: ''),
                  );
                  return pw.TableRow(children: [
                    pw.Text(stockItem.name, style: const pw.TextStyle(fontSize: 8)),
                    pw.Text('${item.qty}${stockItem.unit}', style: const pw.TextStyle(fontSize: 8)),
                    pw.Text(item.rate.toStringAsFixed(0), style: const pw.TextStyle(fontSize: 8)),
                    pw.Text(item.total.toStringAsFixed(0), style: const pw.TextStyle(fontSize: 8), textAlign: pw.TextAlign.right),
                  ]);
                }),
              ],
            ),
            pw.Divider(thickness: 0.5),

            // ٹوٹل
            if (sale.discount > 0)
              pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
                pw.Text('ڈسکاؤنٹ:', style: const pw.TextStyle(fontSize: 8)),
                pw.Text('-Rs. ${sale.discount.toStringAsFixed(2)}', style: const pw.TextStyle(fontSize: 8)),
              ]),
            if (sale.fee > 0)
              pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
                pw.Text('اضافی:', style: const pw.TextStyle(fontSize: 8)),
                pw.Text('+Rs. ${sale.fee.toStringAsFixed(2)}', style: const pw.TextStyle(fontSize: 8)),
              ]),
            pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
              pw.Text('کل:', style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold)),
              pw.Text('Rs. ${sale.total.toStringAsFixed(2)}', style: pw.TextStyle(fontSize: 10, fontWeight: pw.FontWeight.bold)),
            ]),
            pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
              pw.Text('نقد:', style: const pw.TextStyle(fontSize: 8)),
              pw.Text('Rs. ${sale.cashReceived.toStringAsFixed(2)}', style: const pw.TextStyle(fontSize: 8)),
            ]),
            if (sale.creditAmount > 0)
              pw.Row(mainAxisAlignment: pw.MainAxisAlignment.spaceBetween, children: [
                pw.Text('بقایا:', style: const pw.TextStyle(fontSize: 8)),
                pw.Text('Rs. ${sale.creditAmount.toStringAsFixed(2)}', style: const pw.TextStyle(fontSize: 8)),
              ]),
            pw.Divider(thickness: 1),

            // فوٹر
            if (sale.note.isNotEmpty)
              pw.Text('نوٹ: ${sale.note}', style: const pw.TextStyle(fontSize: 7)),
            pw.SizedBox(height: 4),
            pw.Text('شکریہ! دوبارہ تشریف لائیں',
                style: const pw.TextStyle(fontSize: 8),
                textAlign: pw.TextAlign.center),
            pw.SizedBox(height: 8),
          ],
        );
      },
    ));

    await Printing.layoutPdf(
      onLayout: (_) => doc.save(),
      name: 'Sale_Receipt_${sale.id}',
    );
  }

  // ============================================================
  // A4 رپورٹ پرنٹ
  // ============================================================
  static Future<void> printA4Report({
    required String title,
    required String shopName,
    required String dateRange,
    required List<Map<String, String>> rows,
    required Map<String, String> summary,
  }) async {
    final doc = pw.Document();

    doc.addPage(pw.MultiPage(
      pageFormat: PdfPageFormat.a4,
      margin: const pw.EdgeInsets.all(20),
      header: (pw.Context ctx) => pw.Column(
        crossAxisAlignment: pw.CrossAxisAlignment.center,
        children: [
          pw.Text(shopName,
              style: pw.TextStyle(fontSize: 18, fontWeight: pw.FontWeight.bold)),
          pw.Text(title,
              style: pw.TextStyle(fontSize: 14, fontWeight: pw.FontWeight.bold)),
          pw.Text('مدت: $dateRange',
              style: const pw.TextStyle(fontSize: 10)),
          pw.Divider(thickness: 1.5),
        ],
      ),
      footer: (pw.Context ctx) => pw.Row(
        mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
        children: [
          pw.Text('IQBAL TRADERS — خودکار رپورٹ',
              style: const pw.TextStyle(fontSize: 8)),
          pw.Text('صفحہ ${ctx.pageNumber} / ${ctx.pagesCount}',
              style: const pw.TextStyle(fontSize: 8)),
        ],
      ),
      build: (pw.Context ctx) => [
        // سمری باکسز
        if (summary.isNotEmpty) ...[
          pw.Container(
            padding: const pw.EdgeInsets.all(10),
            decoration: pw.BoxDecoration(
              border: pw.Border.all(color: PdfColors.grey),
              borderRadius: const pw.BorderRadius.all(pw.Radius.circular(8)),
            ),
            child: pw.Wrap(
              spacing: 20,
              runSpacing: 10,
              children: summary.entries.map((e) => pw.Column(
                crossAxisAlignment: pw.CrossAxisAlignment.start,
                children: [
                  pw.Text(e.key, style: const pw.TextStyle(fontSize: 9, color: PdfColors.grey700)),
                  pw.Text(e.value, style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
                ],
              )).toList(),
            ),
          ),
          pw.SizedBox(height: 12),
        ],

        // ڈیٹا ٹیبل
        if (rows.isNotEmpty) ...[
          pw.TableHelper.fromTextArray(
            headers: rows.first.keys.toList(),
            data: rows.map((r) => r.values.toList()).toList(),
            headerStyle: pw.TextStyle(fontSize: 9, fontWeight: pw.FontWeight.bold),
            cellStyle: const pw.TextStyle(fontSize: 8),
            cellAlignment: pw.Alignment.centerRight,
            headerDecoration: const pw.BoxDecoration(color: PdfColors.grey200),
            border: pw.TableBorder.all(color: PdfColors.grey300, width: 0.5),
            columnWidths: const {0: pw.FlexColumnWidth(2)},
          ),
        ],
      ],
    ));

    await Printing.layoutPdf(
      onLayout: (_) => doc.save(),
      name: title,
    );
  }

  // ============================================================
  // کسٹمر بیان پرنٹ (Account Statement)
  // ============================================================
  static Future<void> printCustomerStatement({
    required Customer customer,
    required String shopName,
  }) async {
    final doc = pw.Document();

    doc.addPage(pw.MultiPage(
      pageFormat: PdfPageFormat.a4,
      margin: const pw.EdgeInsets.all(20),
      header: (pw.Context ctx) => pw.Column(
        crossAxisAlignment: pw.CrossAxisAlignment.center,
        children: [
          pw.Text(shopName,
              style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold)),
          pw.Text('کسٹمر کھاتہ — ${customer.name}',
              style: pw.TextStyle(fontSize: 13, fontWeight: pw.FontWeight.bold)),
          if (customer.phone.isNotEmpty)
            pw.Text('📞 ${customer.phone}', style: const pw.TextStyle(fontSize: 9)),
          pw.Divider(thickness: 1),
        ],
      ),
      build: (pw.Context ctx) => [
        // موجودہ بیلنس
        pw.Container(
          padding: const pw.EdgeInsets.all(10),
          decoration: pw.BoxDecoration(
            border: pw.Border.all(),
            borderRadius: const pw.BorderRadius.all(pw.Radius.circular(4)),
          ),
          child: pw.Row(
            mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
            children: [
              pw.Text('موجودہ بیلنس:',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold, fontSize: 12)),
              pw.Text('Rs. ${customer.balance.toStringAsFixed(2)}',
                  style: pw.TextStyle(
                      fontWeight: pw.FontWeight.bold,
                      fontSize: 14,
                      color: customer.balance > 0 ? PdfColors.red : PdfColors.green)),
            ],
          ),
        ),
        pw.SizedBox(height: 12),

        // لین دین تاریخچہ
        pw.Text('لین دین تاریخچہ',
            style: pw.TextStyle(fontWeight: pw.FontWeight.bold, fontSize: 12)),
        pw.SizedBox(height: 6),
        pw.TableHelper.fromTextArray(
          headers: ['تاریخ', 'قسم', 'رقم', 'نوٹ'],
          data: customer.transactions.map((t) => [
            t.date,
            t.type,
            'Rs. ${t.amount.abs().toStringAsFixed(2)}',
            t.note ?? '',
          ]).toList(),
          headerStyle: pw.TextStyle(fontSize: 9, fontWeight: pw.FontWeight.bold),
          cellStyle: const pw.TextStyle(fontSize: 8),
          border: pw.TableBorder.all(color: PdfColors.grey300, width: 0.5),
          headerDecoration: const pw.BoxDecoration(color: PdfColors.grey200),
        ),
      ],
    ));

    await Printing.layoutPdf(
      onLayout: (_) => doc.save(),
      name: 'Customer_${customer.name}',
    );
  }

  // ============================================================
  // سپلائر بیان
  // ============================================================
  static Future<void> printSupplierStatement({
    required Supplier supplier,
    required String shopName,
  }) async {
    final doc = pw.Document();

    doc.addPage(pw.MultiPage(
      pageFormat: PdfPageFormat.a4,
      margin: const pw.EdgeInsets.all(20),
      header: (pw.Context ctx) => pw.Column(
        crossAxisAlignment: pw.CrossAxisAlignment.center,
        children: [
          pw.Text(shopName,
              style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold)),
          pw.Text('سپلائر کھاتہ — ${supplier.name}',
              style: pw.TextStyle(fontSize: 13, fontWeight: pw.FontWeight.bold)),
          pw.Divider(thickness: 1),
        ],
      ),
      build: (pw.Context ctx) => [
        pw.Container(
          padding: const pw.EdgeInsets.all(10),
          decoration: pw.BoxDecoration(border: pw.Border.all()),
          child: pw.Row(
            mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
            children: [
              pw.Text('موجودہ بیلنس:', style: pw.TextStyle(fontWeight: pw.FontWeight.bold, fontSize: 12)),
              pw.Text('Rs. ${supplier.balance.toStringAsFixed(2)}',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold, fontSize: 14,
                      color: supplier.balance > 0 ? PdfColors.red : PdfColors.green)),
            ],
          ),
        ),
        pw.SizedBox(height: 12),
        pw.TableHelper.fromTextArray(
          headers: ['تاریخ', 'قسم', 'رقم', 'نوٹ'],
          data: supplier.transactions.map((t) => [
            t.date, t.type,
            'Rs. ${t.amount.abs().toStringAsFixed(2)}',
            t.note ?? '',
          ]).toList(),
          headerStyle: pw.TextStyle(fontSize: 9, fontWeight: pw.FontWeight.bold),
          cellStyle: const pw.TextStyle(fontSize: 8),
          border: pw.TableBorder.all(color: PdfColors.grey300, width: 0.5),
          headerDecoration: const pw.BoxDecoration(color: PdfColors.grey200),
        ),
      ],
    ));

    await Printing.layoutPdf(
      onLayout: (_) => doc.save(),
      name: 'Supplier_${supplier.name}',
    );
  }

  // ============================================================
  // اسٹاک رپورٹ A4
  // ============================================================
  static Future<void> printStockReport({
    required List<StockItem> items,
    required String shopName,
  }) async {
    final doc = pw.Document();
    final totalValue = items.fold(0.0, (s, i) => s + (i.quantity * i.purchaseRate));

    doc.addPage(pw.MultiPage(
      pageFormat: PdfPageFormat.a4,
      margin: const pw.EdgeInsets.all(20),
      header: (pw.Context ctx) => pw.Column(
        crossAxisAlignment: pw.CrossAxisAlignment.center,
        children: [
          pw.Text(shopName,
              style: pw.TextStyle(fontSize: 16, fontWeight: pw.FontWeight.bold)),
          pw.Text('اسٹاک رپورٹ — ${DateTime.now().toIso8601String().split('T')[0]}',
              style: pw.TextStyle(fontSize: 13, fontWeight: pw.FontWeight.bold)),
          pw.Divider(thickness: 1),
        ],
      ),
      build: (pw.Context ctx) => [
        pw.Container(
          padding: const pw.EdgeInsets.all(8),
          decoration: pw.BoxDecoration(border: pw.Border.all()),
          child: pw.Row(
            mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
            children: [
              pw.Text('کل آئٹمز: ${items.length}',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
              pw.Text('کل مالیت: Rs. ${totalValue.toStringAsFixed(2)}',
                  style: pw.TextStyle(fontWeight: pw.FontWeight.bold)),
            ],
          ),
        ),
        pw.SizedBox(height: 10),
        pw.TableHelper.fromTextArray(
          headers: ['آئٹم', 'کیٹیگری', 'مقدار', 'یونٹ', 'WAC ریٹ', 'مالیت', 'الرٹ'],
          data: items.map((i) => [
            i.name,
            i.category,
            i.quantity.toStringAsFixed(2),
            i.unit,
            'Rs. ${i.purchaseRate.toStringAsFixed(2)}',
            'Rs. ${(i.quantity * i.purchaseRate).toStringAsFixed(0)}',
            i.isLowStock ? '⚠️ کم' : '✅',
          ]).toList(),
          headerStyle: pw.TextStyle(fontSize: 8, fontWeight: pw.FontWeight.bold),
          cellStyle: const pw.TextStyle(fontSize: 8),
          border: pw.TableBorder.all(color: PdfColors.grey300, width: 0.5),
          headerDecoration: const pw.BoxDecoration(color: PdfColors.grey200),
        ),
      ],
    ));

    await Printing.layoutPdf(
      onLayout: (_) => doc.save(),
      name: 'Stock_Report',
    );
  }
}


// ============================================================
// IQBAL TRADERS - Reusable Widgets
// ============================================================

// سمری کارڈ (Dashboard)
class SummaryCard extends StatelessWidget {
  final String title;
  final String amount;
  final IconData icon;
  final Color color;
  final Color bgColor;
  final String? subtitle;
  final VoidCallback? onTap;

  const SummaryCard({
    super.key,
    required this.title,
    required this.amount,
    required this.icon,
    required this.color,
    required this.bgColor,
    this.subtitle,
    this.onTap,
  });

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onTap: onTap,
      child: Container(
        padding: const EdgeInsets.all(14),
        decoration: BoxDecoration(
          color: bgColor,
          borderRadius: BorderRadius.circular(12),
          border: Border.all(color: color.withOpacity(0.3)),
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceBetween,
              children: [
                Expanded(
                  child: Text(
                    title,
                    style: TextStyle(
                      color: color,
                      fontWeight: FontWeight.bold,
                      fontSize: 13,
                    ),
                  ),
                ),
                Icon(icon, color: color, size: 22),
              ],
            ),
            const SizedBox(height: 8),
            Text(
              amount,
              style: TextStyle(
                color: color,
                fontWeight: FontWeight.w900,
                fontSize: 18,
              ),
            ),
            if (subtitle != null) ...[
              const SizedBox(height: 4),
              Text(
                subtitle!,
                style: TextStyle(
                  color: color.withOpacity(0.7),
                  fontSize: 11,
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }
}

// گریڈینٹ سمری باکس
class GradientSummaryBox extends StatelessWidget {
  final String title;
  final String amount;
  final IconData icon;
  final List<Color> colors;
  final String? subtitle;
  final VoidCallback? onTap;

  const GradientSummaryBox({
    super.key,
    required this.title,
    required this.amount,
    required this.icon,
    required this.colors,
    this.subtitle,
    this.onTap,
  });

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onTap: onTap,
      child: Container(
        width: double.infinity,
        padding: const EdgeInsets.all(16),
        decoration: BoxDecoration(
          gradient: LinearGradient(
            colors: colors,
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
          ),
          borderRadius: BorderRadius.circular(12),
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Icon(icon, color: Colors.white, size: 22),
                const SizedBox(width: 8),
                Text(
                  title,
                  style: const TextStyle(
                    color: Colors.white,
                    fontWeight: FontWeight.bold,
                    fontSize: 14,
                  ),
                ),
              ],
            ),
            const SizedBox(height: 10),
            Text(
              amount,
              style: const TextStyle(
                color: Colors.white,
                fontWeight: FontWeight.w900,
                fontSize: 24,
              ),
            ),
            if (subtitle != null) ...[
              const SizedBox(height: 4),
              Text(
                subtitle!,
                style: const TextStyle(
                  color: Colors.white70,
                  fontSize: 12,
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }
}

// کارڈ ہیڈر
class AppCardHeader extends StatelessWidget {
  final String title;
  final List<Widget>? actions;
  final IconData? icon;

  const AppCardHeader({
    super.key,
    required this.title,
    this.actions,
    this.icon,
  });

  @override
  Widget build(BuildContext context) {
    return Row(
      children: [
        if (icon != null) ...[
          Icon(icon, color: AppTheme.primary, size: 20),
          const SizedBox(width: 8),
        ],
        Expanded(
          child: Text(
            title,
            style: const TextStyle(
              fontSize: 17,
              fontWeight: FontWeight.bold,
              color: AppTheme.textPrimary,
            ),
          ),
        ),
        if (actions != null) ...actions!,
      ],
    );
  }
}

// اسٹیٹس بیج
class StatusBadge extends StatelessWidget {
  final String label;
  final Color color;

  const StatusBadge({super.key, required this.label, required this.color});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
      decoration: BoxDecoration(
        color: color.withOpacity(0.15),
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: color.withOpacity(0.3)),
      ),
      child: Text(
        label,
        style: TextStyle(
          color: color,
          fontSize: 11,
          fontWeight: FontWeight.bold,
        ),
      ),
    );
  }
}

// بٹن گروپ
class ActionButtonRow extends StatelessWidget {
  final List<_ActionButtonData> buttons;

  const ActionButtonRow({super.key, required this.buttons});

  @override
  Widget build(BuildContext context) {
    return Wrap(
      spacing: 8,
      runSpacing: 8,
      children: buttons
          .map((b) => ElevatedButton.icon(
                onPressed: b.onPressed,
                icon: Icon(b.icon, size: 16),
                label: Text(b.label, style: const TextStyle(fontSize: 13)),
                style: ElevatedButton.styleFrom(
                  backgroundColor: b.color,
                  foregroundColor: Colors.white,
                  padding:
                      const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                  shape: RoundedRectangleBorder(
                      borderRadius: BorderRadius.circular(8)),
                ),
              ))
          .toList(),
    );
  }
}

class _ActionButtonData {
  final String label;
  final IconData icon;
  final Color color;
  final VoidCallback onPressed;

  _ActionButtonData({
    required this.label,
    required this.icon,
    required this.color,
    required this.onPressed,
  });
}

ActionButtonRow buildActionButtons(List<_ActionButtonData> buttons) =>
    ActionButtonRow(buttons: buttons);

// سرچ باکس
class AppSearchBox extends StatelessWidget {
  final TextEditingController controller;
  final String hint;
  final ValueChanged<String> onChanged;

  const AppSearchBox({
    super.key,
    required this.controller,
    required this.hint,
    required this.onChanged,
  });

  @override
  Widget build(BuildContext context) {
    return TextField(
      controller: controller,
      onChanged: onChanged,
      textAlign: TextAlign.right,
      decoration: InputDecoration(
        hintText: hint,
        hintStyle: const TextStyle(fontSize: 14),
        prefixIcon: const Icon(Icons.search, color: AppTheme.primary),
        suffixIcon: controller.text.isNotEmpty
            ? IconButton(
                icon: const Icon(Icons.clear),
                onPressed: () {
                  controller.clear();
                  onChanged('');
                },
              )
            : null,
        filled: true,
        fillColor: AppTheme.lightBg,
        border: OutlineInputBorder(
          borderRadius: BorderRadius.circular(10),
          borderSide: const BorderSide(color: AppTheme.cardBorder),
        ),
        focusedBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(10),
          borderSide: const BorderSide(color: AppTheme.primary, width: 2),
        ),
        contentPadding:
            const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
      ),
    );
  }
}

// خالی حالت
class EmptyState extends StatelessWidget {
  final String message;
  final IconData icon;
  final String? buttonLabel;
  final VoidCallback? onButtonPress;

  const EmptyState({
    super.key,
    required this.message,
    required this.icon,
    this.buttonLabel,
    this.onButtonPress,
  });

  @override
  Widget build(BuildContext context) {
    return Center(
      child: Padding(
        padding: const EdgeInsets.all(40),
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            Icon(icon, size: 64, color: Colors.grey.shade400),
            const SizedBox(height: 16),
            Text(
              message,
              textAlign: TextAlign.center,
              style: TextStyle(
                  fontSize: 16,
                  color: Colors.grey.shade600,
                  fontWeight: FontWeight.bold),
            ),
            if (buttonLabel != null && onButtonPress != null) ...[
              const SizedBox(height: 20),
              ElevatedButton.icon(
                onPressed: onButtonPress,
                icon: const Icon(Icons.add),
                label: Text(buttonLabel!),
              ),
            ],
          ],
        ),
      ),
    );
  }
}

// کنفرم ڈائیلاگ
Future<bool> showConfirmDialog(
    BuildContext context, String title, String message,
    {Color? confirmColor}) async {
  return await showDialog<bool>(
        context: context,
        builder: (ctx) => AlertDialog(
          title: Text(title,
              style: const TextStyle(fontWeight: FontWeight.bold)),
          content: Text(message),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(ctx, false),
                child: const Text('نہیں')),
            ElevatedButton(
              onPressed: () => Navigator.pop(ctx, true),
              style: ElevatedButton.styleFrom(
                  backgroundColor: confirmColor ?? AppTheme.danger,
                  foregroundColor: Colors.white),
              child: const Text('ہاں'),
            ),
          ],
        ),
      ) ??
      false;
}

// اسنیک بار
void showSnackBar(BuildContext context, String message,
    {bool isError = false}) {
  ScaffoldMessenger.of(context).showSnackBar(SnackBar(
    content: Text(message),
    backgroundColor: isError ? AppTheme.danger : AppTheme.success,
    behavior: SnackBarBehavior.floating,
    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
  ));
}

// لوڈنگ انڈیکیٹر
class LoadingOverlay extends StatelessWidget {
  const LoadingOverlay({super.key});

  @override
  Widget build(BuildContext context) {
    return Container(
      color: Colors.black26,
      child: const Center(
        child: CircularProgressIndicator(color: AppTheme.primary),
      ),
    );
  }
}

// ٹرانزیکشن ٹائل
class TransactionTile extends StatelessWidget {
  final String date;
  final String type;
  final double amount;
  final String? note;
  final bool isPositive;
  final VoidCallback? onDelete;
  final VoidCallback? onEdit;

  const TransactionTile({
    super.key,
    required this.date,
    required this.type,
    required this.amount,
    this.note,
    required this.isPositive,
    this.onDelete,
    this.onEdit,
  });

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
      decoration: BoxDecoration(
        color: isPositive
            ? AppTheme.success.withOpacity(0.05)
            : AppTheme.danger.withOpacity(0.05),
        border: Border(
          bottom: BorderSide(color: AppTheme.cardBorder.withOpacity(0.5)),
          right: BorderSide(
              color: isPositive ? AppTheme.success : AppTheme.danger,
              width: 3),
        ),
      ),
      child: Row(
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Text(
                      type,
                      style: const TextStyle(
                          fontWeight: FontWeight.bold, fontSize: 14),
                    ),
                    const SizedBox(width: 8),
                    Text(
                      date,
                      style: TextStyle(
                          fontSize: 11, color: Colors.grey.shade600),
                    ),
                  ],
                ),
                if (note != null && note!.isNotEmpty) ...[
                  const SizedBox(height: 2),
                  Text(
                    note!,
                    style: TextStyle(
                        fontSize: 12, color: Colors.grey.shade500),
                  ),
                ],
              ],
            ),
          ),
          Text(
            '${isPositive ? '+' : '-'}Rs. ${amount.abs().toStringAsFixed(2)}',
            style: TextStyle(
              color: isPositive ? AppTheme.success : AppTheme.danger,
              fontWeight: FontWeight.w900,
              fontSize: 14,
            ),
          ),
          if (onDelete != null || onEdit != null) ...[
            const SizedBox(width: 8),
            if (onEdit != null)
              IconButton(
                icon: const Icon(Icons.edit,
                    size: 16, color: AppTheme.orange),
                onPressed: onEdit,
                padding: EdgeInsets.zero,
                constraints: const BoxConstraints(),
              ),
            if (onDelete != null)
              IconButton(
                icon: const Icon(Icons.delete,
                    size: 16, color: AppTheme.danger),
                onPressed: onDelete,
                padding: EdgeInsets.zero,
                constraints: const BoxConstraints(),
              ),
          ],
        ],
      ),
    );
  }
}

// فارم فیلڈ
class AppFormField extends StatelessWidget {
  final String label;
  final TextEditingController controller;
  final TextInputType keyboardType;
  final String? hint;
  final bool required;
  final String? Function(String?)? validator;
  final int maxLines;
  final bool readOnly;
  final VoidCallback? onTap;
  final Widget? suffix;

  const AppFormField({
    super.key,
    required this.label,
    required this.controller,
    this.keyboardType = TextInputType.text,
    this.hint,
    this.required = false,
    this.validator,
    this.maxLines = 1,
    this.readOnly = false,
    this.onTap,
    this.suffix,
  });

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(
          '$label${required ? ' *' : ''}',
          style: const TextStyle(
              fontWeight: FontWeight.bold,
              fontSize: 14,
              color: AppTheme.textPrimary),
        ),
        const SizedBox(height: 6),
        TextFormField(
          controller: controller,
          keyboardType: keyboardType,
          textAlign: TextAlign.right,
          maxLines: maxLines,
          readOnly: readOnly,
          onTap: onTap,
          validator: validator ??
              (required
                  ? (v) =>
                      (v == null || v.trim().isEmpty) ? '$label ضروری ہے' : null
                  : null),
          decoration: InputDecoration(
            hintText: hint,
            hintStyle: const TextStyle(fontSize: 13),
            suffixIcon: suffix,
          ),
        ),
      ],
    );
  }
}

// ڈیٹ پکر فیلڈ
class DatePickerField extends StatelessWidget {
  final String label;
  final TextEditingController controller;
  final bool required;

  const DatePickerField({
    super.key,
    required this.label,
    required this.controller,
    this.required = true,
  });

  @override
  Widget build(BuildContext context) {
    return AppFormField(
      label: label,
      controller: controller,
      required: required,
      readOnly: true,
      onTap: () async {
        final picked = await showDatePicker(
          context: context,
          initialDate: DateTime.now(),
          firstDate: DateTime(2000),
          lastDate: DateTime(2100),
          builder: (ctx, child) => Theme(
            data: Theme.of(ctx).copyWith(
              colorScheme: const ColorScheme.light(
                  primary: AppTheme.primary),
            ),
            child: child!,
          ),
        );
        if (picked != null) {
          controller.text =
              picked.toIso8601String().split('T')[0];
        }
      },
      suffix: const Icon(Icons.calendar_today, color: AppTheme.primary),
    );
  }
}

// بیلنس ڈسپلے
class BalanceDisplay extends StatelessWidget {
  final double balance;
  final String positiveLabel;
  final String negativeLabel;
  final String zeroLabel;

  const BalanceDisplay({
    super.key,
    required this.balance,
    this.positiveLabel = 'واجب الادا',
    this.negativeLabel = 'اضافی ادائیگی',
    this.zeroLabel = 'برابر',
  });

  @override
  Widget build(BuildContext context) {
    Color color;
    String label;
    if (balance > 0) {
      color = AppTheme.danger;
      label = positiveLabel;
    } else if (balance < 0) {
      color = AppTheme.success;
      label = negativeLabel;
    } else {
      color = Colors.grey;
      label = zeroLabel;
    }
    return Column(
      children: [
        Text(
          'Rs. ${balance.abs().toStringAsFixed(2)}',
          style: TextStyle(
            color: color,
            fontWeight: FontWeight.w900,
            fontSize: 28,
          ),
        ),
        const SizedBox(height: 4),
        StatusBadge(label: label, color: color),
      ],
    );
  }
}

// آئٹم سیلیکشن کارڈ (فروخت/خریداری)
class ItemSelectionCard extends StatelessWidget {
  final int itemId;
  final String itemName;
  final String unit;
  final double availableQty;
  final double rate;
  final double qty;
  final double total;
  final ValueChanged<double> onQtyChanged;
  final ValueChanged<double> onRateChanged;
  final VoidCallback onRemove;

  const ItemSelectionCard({
    super.key,
    required this.itemId,
    required this.itemName,
    required this.unit,
    required this.availableQty,
    required this.rate,
    required this.qty,
    required this.total,
    required this.onQtyChanged,
    required this.onRateChanged,
    required this.onRemove,
  });

  @override
  Widget build(BuildContext context) {
    return Card(
      margin: const EdgeInsets.only(bottom: 10),
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Expanded(
                  child: Text(
                    itemName,
                    style: const TextStyle(
                        fontWeight: FontWeight.bold, fontSize: 15),
                  ),
                ),
                Text(
                  'دستیاب: $availableQty $unit',
                  style: TextStyle(
                      fontSize: 11, color: Colors.grey.shade600),
                ),
                const SizedBox(width: 8),
                IconButton(
                  icon: const Icon(Icons.delete, color: AppTheme.danger),
                  onPressed: onRemove,
                  padding: EdgeInsets.zero,
                  constraints: const BoxConstraints(),
                ),
              ],
            ),
            const SizedBox(height: 8),
            Row(
              children: [
                Expanded(
                  child: _miniField(
                    label: 'مقدار ($unit)',
                    initialValue: qty.toString(),
                    onChanged: (v) =>
                        onQtyChanged(double.tryParse(v) ?? 0),
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: _miniField(
                    label: 'ریٹ (Rs.)',
                    initialValue: rate.toStringAsFixed(2),
                    onChanged: (v) =>
                        onRateChanged(double.tryParse(v) ?? 0),
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Text('کل',
                          style: TextStyle(
                              fontSize: 11,
                              fontWeight: FontWeight.bold)),
                      Container(
                        padding: const EdgeInsets.symmetric(
                            horizontal: 8, vertical: 10),
                        decoration: BoxDecoration(
                          color: AppTheme.primary.withOpacity(0.1),
                          borderRadius: BorderRadius.circular(6),
                        ),
                        child: Text(
                          'Rs. ${total.toStringAsFixed(2)}',
                          style: const TextStyle(
                              fontWeight: FontWeight.w900,
                              color: AppTheme.primary,
                              fontSize: 13),
                        ),
                      ),
                    ],
                  ),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Widget _miniField(
      {required String label,
      required String initialValue,
      required ValueChanged<String> onChanged}) {
    final controller = TextEditingController(text: initialValue);
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(label,
            style: const TextStyle(
                fontSize: 11, fontWeight: FontWeight.bold)),
        const SizedBox(height: 4),
        TextFormField(
          controller: controller,
          keyboardType:
              const TextInputType.numberWithOptions(decimal: true),
          textAlign: TextAlign.right,
          onChanged: onChanged,
          decoration: InputDecoration(
            contentPadding:
                const EdgeInsets.symmetric(horizontal: 8, vertical: 8),
            border: OutlineInputBorder(
                borderRadius: BorderRadius.circular(6)),
            isDense: true,
          ),
          style: const TextStyle(fontWeight: FontWeight.bold),
        ),
      ],
    );
  }
}


class LoadingWidget extends StatelessWidget {
  final String? message;
  const LoadingWidget({super.key, this.message});

  @override
  Widget build(BuildContext context) {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          const CircularProgressIndicator(color: AppTheme.primary),
          if (message != null) ...[
            const SizedBox(height: 16),
            Text(message!, style: const TextStyle(color: AppTheme.textPrimary)),
          ],
        ],
      ),
    );
  }
}


class AppTextField extends StatelessWidget {
  final TextEditingController controller;
  final String label;
  final String? hint;
  final TextInputType keyboardType;
  final List<TextInputFormatter>? inputFormatters;
  final String? Function(String?)? validator;
  final int maxLines;
  final bool readOnly;
  final VoidCallback? onTap;
  final Widget? suffixIcon;
  final Widget? prefixIcon;

  const AppTextField({
    super.key,
    required this.controller,
    required this.label,
    this.hint,
    this.keyboardType = TextInputType.text,
    this.inputFormatters,
    this.validator,
    this.maxLines = 1,
    this.readOnly = false,
    this.onTap,
    this.suffixIcon,
    this.prefixIcon,
  });

  @override
  Widget build(BuildContext context) {
    return TextFormField(
      controller: controller,
      keyboardType: keyboardType,
      inputFormatters: inputFormatters,
      validator: validator,
      maxLines: maxLines,
      readOnly: readOnly,
      onTap: onTap,
      decoration: InputDecoration(
        labelText: label,
        hintText: hint,
        suffixIcon: suffixIcon,
        prefixIcon: prefixIcon,
      ),
    );
  }
}


// ============================================================
// APP WIDGET
// ============================================================

class IqbalTradersApp extends StatelessWidget {
  const IqbalTradersApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'IQBAL TRADERS',
      debugShowCheckedModeBanner: false,
      theme: AppTheme.theme,
      locale: const Locale('ur', 'PK'),
      supportedLocales: const [Locale('ur', 'PK'), Locale('en', 'US')],
      localizationsDelegates: const [
        GlobalMaterialLocalizations.delegate,
        GlobalWidgetsLocalizations.delegate,
        GlobalCupertinoLocalizations.delegate,
      ],
      initialRoute: '/login',
      routes: AppRoutes.routes,
      builder: (context, child) {
        return Consumer<BusinessProvider>(
          builder: (context, bp, _) {
            if (bp.isLoading) return const _SplashScreen();
            return child ?? const SizedBox.shrink();
          },
        );
      },
    );
  }
}

class _SplashScreen extends StatelessWidget {
  const _SplashScreen();
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        decoration: const BoxDecoration(
          gradient: LinearGradient(
            colors: [AppTheme.primary, AppTheme.primaryDark],
            begin: Alignment.topLeft, end: Alignment.bottomRight,
          ),
        ),
        child: Center(
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Container(
                padding: const EdgeInsets.all(24),
                decoration: BoxDecoration(
                  color: Colors.white.withOpacity(0.15), shape: BoxShape.circle,
                ),
                child: const Icon(Icons.store_rounded, size: 72, color: Colors.white),
              ),
              const SizedBox(height: 24),
              const Text('IQBAL TRADERS',
                style: TextStyle(color: Colors.white, fontSize: 30, fontWeight: FontWeight.w900, letterSpacing: 2)),
              const SizedBox(height: 8),
              const Text('کاروباری مینجمنٹ سسٹم',
                style: TextStyle(color: Colors.white70, fontSize: 16)),
              const SizedBox(height: 48),
              const CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
              const SizedBox(height: 16),
              const Text('لوڈ ہو رہا ہے...', style: TextStyle(color: Colors.white60, fontSize: 13)),
            ],
          ),
        ),
      ),
    );
  }
}

// ============================================================
// ROUTES
// ============================================================

class AppRoutes {
  static const login        = '/login';
  static const dashboard    = '/dashboard';
  static const customers    = '/customers';
  static const suppliers    = '/suppliers';
  static const udhar        = '/udhar';
  static const stock        = '/stock';
  static const transactions = '/transactions';
  static const addSale      = '/add-sale';
  static const expenses     = '/expenses';
  static const reports      = '/reports';
  static const profit       = '/profit';
  static const profitCash   = '/profit-cash';
  static const cashInOut    = '/cash-in-out';
  static const partners     = '/partners';
  static const settings     = '/settings';

  static Map<String, WidgetBuilder> get routes => {
    login:        (_) => const LoginScreen(),
    dashboard:    (_) => const MainNavigationScreen(),
    customers:    (_) => const CustomersScreen(),
    suppliers:    (_) => const SuppliersScreen(),
    udhar:        (_) => const UdharScreen(),
    stock:        (_) => const StockScreen(),
    transactions: (_) => const TransactionsScreen(),
    expenses:     (_) => const ExpensesScreen(),
    reports:      (_) => const ReportsScreen(),
    profit:       (_) => const ProfitScreen(),
    profitCash:   (_) => const ProfitScreen(),
    cashInOut:    (_) => const CashInOutScreen(),
    partners:     (_) => const PartnersScreen(),
    settings:     (_) => const SettingsScreen(),
  };
}


// ============================================================
// لاگ ان اسکرین
// ============================================================
class LoginScreen extends StatefulWidget {
  const LoginScreen({super.key});

  @override
  State<LoginScreen> createState() => _LoginScreenState();
}

class _LoginScreenState extends State<LoginScreen> {
  final _passwordCtrl = TextEditingController();
  bool _showPassword = false;
  bool _isLoading = false;
  String? _error;

  @override
  void dispose() {
    _passwordCtrl.dispose();
    super.dispose();
  }

  Future<void> _login() async {
    final pass = _passwordCtrl.text.trim();
    if (pass.isEmpty) {
      setState(() => _error = '❌ پاسورڈ درج کریں!');
      return;
    }
    setState(() {
      _isLoading = true;
      _error = null;
    });
    final provider = context.read<BusinessProvider>();
    final valid = await provider.verifyPassword(pass);
    setState(() => _isLoading = false);
    if (valid) {
      if (mounted) {
        Navigator.pushReplacementNamed(context, '/dashboard');
      }
    } else {
      setState(() => _error = '❌ غلط پاسورڈ! دوبارہ کوشش کریں۔');
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        decoration: const BoxDecoration(
          gradient: LinearGradient(
            colors: [AppTheme.primary, AppTheme.primaryDark],
            begin: Alignment.topLeft,
            end: Alignment.bottomRight,
          ),
        ),
        child: SafeArea(
          child: Center(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24),
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  // لوگو
                  Container(
                    padding: const EdgeInsets.all(20),
                    decoration: BoxDecoration(
                      color: Colors.white.withOpacity(0.15),
                      shape: BoxShape.circle,
                    ),
                    child: const Icon(
                      Icons.store,
                      size: 60,
                      color: Colors.white,
                    ),
                  ),
                  const SizedBox(height: 20),
                  const Text(
                    'IQBAL TRADERS',
                    style: TextStyle(
                      color: Colors.white,
                      fontSize: 28,
                      fontWeight: FontWeight.w900,
                      letterSpacing: 1,
                    ),
                  ),
                  const SizedBox(height: 8),
                  const Text(
                    'کاروباری مینجمنٹ سسٹم',
                    style: TextStyle(
                      color: Colors.white70,
                      fontSize: 16,
                    ),
                  ),
                  const SizedBox(height: 40),

                  // لاگ ان باکس
                  Container(
                    padding: const EdgeInsets.all(28),
                    decoration: BoxDecoration(
                      color: Colors.white,
                      borderRadius: BorderRadius.circular(20),
                      boxShadow: [
                        BoxShadow(
                          color: Colors.black.withOpacity(0.2),
                          blurRadius: 20,
                          offset: const Offset(0, 10),
                        ),
                      ],
                    ),
                    child: Column(
                      children: [
                        const Text(
                          '🔐 داخل ہوں',
                          style: TextStyle(
                            fontSize: 22,
                            fontWeight: FontWeight.bold,
                            color: AppTheme.primary,
                          ),
                        ),
                        const SizedBox(height: 24),
                        TextFormField(
                          controller: _passwordCtrl,
                          obscureText: !_showPassword,
                          textAlign: TextAlign.right,
                          onFieldSubmitted: (_) => _login(),
                          keyboardType: TextInputType.number,
                          decoration: InputDecoration(
                            labelText: '🔑 پاسورڈ',
                            hintText: 'پاسورڈ یہاں درج کریں',
                            prefixIcon: IconButton(
                              icon: Icon(
                                _showPassword
                                    ? Icons.visibility_off
                                    : Icons.visibility,
                                color: AppTheme.primary,
                              ),
                              onPressed: () => setState(
                                  () => _showPassword = !_showPassword),
                            ),
                            border: OutlineInputBorder(
                              borderRadius: BorderRadius.circular(12),
                            ),
                          ),
                        ),
                        if (_error != null) ...[
                          const SizedBox(height: 12),
                          Container(
                            padding: const EdgeInsets.all(10),
                            decoration: BoxDecoration(
                              color: AppTheme.danger.withOpacity(0.1),
                              borderRadius: BorderRadius.circular(8),
                              border: Border.all(
                                  color: AppTheme.danger.withOpacity(0.3)),
                            ),
                            child: Text(
                              _error!,
                              style: const TextStyle(
                                  color: AppTheme.danger,
                                  fontWeight: FontWeight.bold),
                              textAlign: TextAlign.center,
                            ),
                          ),
                        ],
                        const SizedBox(height: 20),
                        SizedBox(
                          width: double.infinity,
                          height: 50,
                          child: ElevatedButton(
                            onPressed: _isLoading ? null : _login,
                            style: ElevatedButton.styleFrom(
                              backgroundColor: AppTheme.primary,
                              foregroundColor: Colors.white,
                              shape: RoundedRectangleBorder(
                                  borderRadius: BorderRadius.circular(12)),
                            ),
                            child: _isLoading
                                ? const CircularProgressIndicator(
                                    color: Colors.white)
                                : const Text(
                                    '🚀 داخل ہوں',
                                    style: TextStyle(
                                        fontSize: 18,
                                        fontWeight: FontWeight.bold),
                                  ),
                          ),
                        ),
                        const SizedBox(height: 12),
                        Text(
                          'پہلی بار؟ ڈیفالٹ پاسورڈ: 1234',
                          style: TextStyle(
                              fontSize: 12, color: Colors.grey.shade500),
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(height: 30),
                  Text(
                    '© ${DateTime.now().year} IQBAL TRADERS',
                    style: const TextStyle(
                        color: Colors.white60, fontSize: 12),
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}


// ============================================================
// مین نیویگیشن اسکرین — ڈرائر + بٹم نیو
// ============================================================

class MainNavigationScreen extends StatefulWidget {
  const MainNavigationScreen({super.key});

  @override
  State<MainNavigationScreen> createState() => _MainNavigationScreenState();
}

class _MainNavigationScreenState extends State<MainNavigationScreen> {
  int _currentIndex = 0;

  final List<_NavItem> _navItems = const [
    _NavItem(screen: DashboardScreen(), label: 'ڈیش بورڈ', icon: Icons.dashboard),
    _NavItem(screen: TransactionsScreen(), label: 'ٹرانزیکشن', icon: Icons.swap_horiz),
    _NavItem(screen: StockScreen(), label: 'اسٹاک', icon: Icons.inventory_2),
    _NavItem(screen: ReportsScreen(), label: 'رپورٹس', icon: Icons.bar_chart),
  ];

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: IndexedStack(
        index: _currentIndex,
        children: _navItems.map((i) => i.screen).toList(),
      ),
      bottomNavigationBar: BottomNavigationBar(
        currentIndex: _currentIndex,
        onTap: (i) => setState(() => _currentIndex = i),
        type: BottomNavigationBarType.fixed,
        selectedItemColor: AppTheme.primary,
        unselectedItemColor: Colors.grey,
        selectedFontSize: 11,
        unselectedFontSize: 10,
        items: _navItems
            .map((i) => BottomNavigationBarItem(icon: Icon(i.icon), label: i.label))
            .toList(),
      ),
      drawer: _buildDrawer(context),
    );
  }

  Widget _buildDrawer(BuildContext context) {
    return Drawer(
      child: ListView(
        padding: EdgeInsets.zero,
        children: [
          DrawerHeader(
            decoration: const BoxDecoration(
              gradient: LinearGradient(
                colors: [AppTheme.primary, AppTheme.primaryDark],
                begin: Alignment.topLeft,
                end: Alignment.bottomRight,
              ),
            ),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              mainAxisAlignment: MainAxisAlignment.end,
              children: [
                const CircleAvatar(
                  radius: 28,
                  backgroundColor: Colors.white24,
                  child: Icon(Icons.store, color: Colors.white, size: 30),
                ),
                const SizedBox(height: 10),
                const Text('IQBAL TRADERS',
                    style: TextStyle(color: Colors.white, fontSize: 18, fontWeight: FontWeight.bold)),
                Text(DateTime.now().toIso8601String().split('T')[0],
                    style: const TextStyle(color: Colors.white70, fontSize: 12)),
              ],
            ),
          ),

          // مینیو آئٹمز
          _drawerItem(context, Icons.dashboard, 'ڈیش بورڈ', AppTheme.primary,
              () { _closeAndNavigate(context); setState(() => _currentIndex = 0); }),
          _drawerItem(context, Icons.swap_horiz, 'ٹرانزیکشنز', const Color(0xFF343a40),
              () { _closeAndNavigate(context); setState(() => _currentIndex = 1); }),
          _drawerItem(context, Icons.inventory_2, 'اسٹاک', AppTheme.info,
              () { _closeAndNavigate(context); setState(() => _currentIndex = 2); }),
          _drawerItem(context, Icons.bar_chart, 'رپورٹس', AppTheme.purple,
              () { _closeAndNavigate(context); setState(() => _currentIndex = 3); }),

          const Divider(),

          _drawerItem(context, Icons.people, 'کسٹمرز', AppTheme.primary,
              () => _navigate(context, const CustomersScreen())),
          _drawerItem(context, Icons.local_shipping, 'سپلائرز', AppTheme.warning,
              () => _navigate(context, const SuppliersScreen())),
          _drawerItem(context, Icons.handshake, 'ادھار', AppTheme.purple,
              () => _navigate(context, const UdharScreen())),
          _drawerItem(context, Icons.money_off, 'اخراجات', AppTheme.danger,
              () => _navigate(context, const ExpensesScreen())),
          _drawerItem(context, Icons.trending_up, 'منافع تجزیہ', AppTheme.success,
              () => _navigate(context, const ProfitScreen())),
          _drawerItem(context, Icons.account_balance_wallet, 'کیش ان/آؤٹ', AppTheme.primary,
              () => _navigate(context, const CashInOutScreen())),
          _drawerItem(context, Icons.group, 'پارٹنرز', AppTheme.purple,
              () => _navigate(context, const PartnersScreen())),

          const Divider(),

          _drawerItem(context, Icons.settings, 'سیٹنگز', const Color(0xFF343a40),
              () => _navigate(context, const SettingsScreen())),

          const Divider(),

          ListTile(
            leading: const Icon(Icons.logout, color: AppTheme.danger),
            title: const Text('🚪 لاگ آؤٹ',
                style: TextStyle(color: AppTheme.danger, fontWeight: FontWeight.bold)),
            onTap: () {
              Navigator.pop(context);
              Navigator.pushNamedAndRemoveUntil(context, '/login', (_) => false);
            },
          ),
        ],
      ),
    );
  }

  Widget _drawerItem(
      BuildContext context, IconData icon, String label, Color color, VoidCallback onTap) {
    return ListTile(
      leading: Icon(icon, color: color, size: 22),
      title: Text(label, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
      onTap: onTap,
    );
  }

  void _closeAndNavigate(BuildContext context) {
    if (Navigator.canPop(context)) Navigator.pop(context);
  }

  void _navigate(BuildContext context, Widget screen) {
    Navigator.pop(context);
    Navigator.push(context, MaterialPageRoute(builder: (_) => screen));
  }
}

class _NavItem {
  final Widget screen;
  final String label;
  final IconData icon;
  const _NavItem({required this.screen, required this.label, required this.icon});
}



// ============================================================
// ڈیش بورڈ اسکرین — مکمل سمری اور آج کی ٹرانزیکشنز
// ============================================================

class DashboardScreen extends StatefulWidget {
  const DashboardScreen({super.key});

  @override
  State<DashboardScreen> createState() => _DashboardScreenState();
}

class _DashboardScreenState extends State<DashboardScreen> {
  String _today() => DateTime.now().toIso8601String().split('T')[0];
  int _selectedTabIndex = 0;

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(
      builder: (context, bp, _) {
        return Scaffold(
          backgroundColor: AppTheme.lightBg,
          body: bp.isLoading
              ? const Center(
                  child: CircularProgressIndicator(color: AppTheme.primary))
              : RefreshIndicator(
                  onRefresh: () => bp.reload(),
                  child: CustomScrollView(
                    slivers: [
                      // AppBar
                      SliverAppBar(
                        expandedHeight: 160,
                        floating: false,
                        pinned: true,
                        flexibleSpace: FlexibleSpaceBar(
                          background: Container(
                            decoration: const BoxDecoration(
                              gradient: LinearGradient(
                                colors: [
                                  AppTheme.primary,
                                  AppTheme.primaryDark
                                ],
                                begin: Alignment.topLeft,
                                end: Alignment.bottomRight,
                              ),
                            ),
                            child: SafeArea(
                              child: Padding(
                                padding: const EdgeInsets.all(16),
                                child: Column(
                                  crossAxisAlignment:
                                      CrossAxisAlignment.center,
                                  mainAxisAlignment: MainAxisAlignment.end,
                                  children: [
                                    Text(
                                      bp.shopName,
                                      style: const TextStyle(
                                        color: Colors.white,
                                        fontSize: 22,
                                        fontWeight: FontWeight.w900,
                                      ),
                                    ),
                                    const SizedBox(height: 6),
                                    Text(
                                      '📅 ${DateTime.now().toString().split(' ')[0]}',
                                      style: const TextStyle(
                                          color: Colors.white70,
                                          fontSize: 13),
                                    ),
                                  ],
                                ),
                              ),
                            ),
                          ),
                        ),
                        actions: [
                          IconButton(
                            icon: const Icon(Icons.undo, color: Colors.white),
                            tooltip: 'Undo',
                            onPressed: () async {
                              final ok = await bp.undo();
                              if (mounted) {
                                showSnackBar(context,
                                    ok ? '✅ Undo کامیاب!' : '❌ Undo نہیں ہو سکا',
                                    isError: !ok);
                              }
                            },
                          ),
                          IconButton(
                            icon: const Icon(Icons.settings,
                                color: Colors.white),
                            onPressed: () =>
                                Navigator.pushNamed(context, '/settings'),
                          ),
                        ],
                      ),

                      SliverToBoxAdapter(
                        child: Padding(
                          padding: const EdgeInsets.all(14),
                          child: Column(
                            children: [
                              // کیش ان ہینڈ بڑا باکس
                              GradientSummaryBox(
                                title: '💰 کیش ان ہینڈ',
                                amount:
                                    'Rs. ${bp.cashInHand.toStringAsFixed(2)}',
                                icon: Icons.account_balance_wallet,
                                colors: const [
                                  AppTheme.primary,
                                  AppTheme.primaryDark
                                ],
                                subtitle:
                                    'آج وصولی: Rs. ${bp.todayCashReceived.toStringAsFixed(2)} | آج ادائیگی: Rs. ${bp.todayCashPaid.toStringAsFixed(2)}',
                              ),
                              const SizedBox(height: 12),

                              // سمری گرڈ — 2 کالم
                              GridView.count(
                                shrinkWrap: true,
                                physics:
                                    const NeverScrollableScrollPhysics(),
                                crossAxisCount: 2,
                                crossAxisSpacing: 10,
                                mainAxisSpacing: 10,
                                childAspectRatio: 1.3,
                                children: [
                                  SummaryCard(
                                    title: '📈 کل خالص منافع',
                                    amount:
                                        'Rs. ${bp.netProfit.toStringAsFixed(2)}',
                                    icon: Icons.trending_up,
                                    color: bp.netProfit >= 0
                                        ? AppTheme.success
                                        : AppTheme.danger,
                                    bgColor: bp.netProfit >= 0
                                        ? AppTheme.success.withOpacity(0.08)
                                        : AppTheme.danger.withOpacity(0.08),
                                    onTap: () => Navigator.pushNamed(
                                        context, '/profit'),
                                  ),
                                  SummaryCard(
                                    title: '📦 اسٹاک مالیت',
                                    amount:
                                        'Rs. ${bp.totalStockValue.toStringAsFixed(2)}',
                                    icon: Icons.inventory_2,
                                    color: AppTheme.info,
                                    bgColor:
                                        AppTheme.info.withOpacity(0.08),
                                    subtitle:
                                        '⚠️ کم اسٹاک: ${bp.lowStockCount}',
                                    onTap: () => Navigator.pushNamed(
                                        context, '/stock'),
                                  ),
                                  SummaryCard(
                                    title: '👥 کسٹمر واجبات',
                                    amount:
                                        'Rs. ${bp.totalReceivables.toStringAsFixed(2)}',
                                    icon: Icons.people,
                                    color: AppTheme.danger,
                                    bgColor:
                                        AppTheme.danger.withOpacity(0.08),
                                    onTap: () => Navigator.pushNamed(
                                        context, '/customers'),
                                  ),
                                  SummaryCard(
                                    title: '🚚 سپلائر واجبات',
                                    amount:
                                        'Rs. ${bp.totalPayables.toStringAsFixed(2)}',
                                    icon: Icons.local_shipping,
                                    color: AppTheme.warning,
                                    bgColor:
                                        AppTheme.warning.withOpacity(0.08),
                                    onTap: () => Navigator.pushNamed(
                                        context, '/suppliers'),
                                  ),
                                ],
                              ),
                              const SizedBox(height: 12),

                              // آج کی سمری
                              _buildTodaySummary(bp),
                              const SizedBox(height: 12),

                              // آج کی ٹرانزیکشنز
                              _buildTodayTransactions(bp),
                              const SizedBox(height: 12),

                              // اسٹاک لسٹ
                              _buildStockSummary(bp),
                              const SizedBox(height: 80),
                            ],
                          ),
                        ),
                      ),
                    ],
                  ),
                ),

          // فلوٹنگ ایکشن بٹن — فوری فروخت
          floatingActionButton: FloatingActionButton.extended(
            onPressed: () => Navigator.pushNamed(context, '/transactions'),
            backgroundColor: AppTheme.success,
            icon: const Icon(Icons.add_shopping_cart, color: Colors.white),
            label: const Text(
              'نئی فروخت',
              style: TextStyle(
                  color: Colors.white, fontWeight: FontWeight.bold),
            ),
          ),
        );
      },
    );
  }

  // آج کی سمری بار
  Widget _buildTodaySummary(BusinessProvider bp) {
    final adj = ({'loss':0.0,'gain':0.0});
    return Container(
      padding: const EdgeInsets.all(14),
      decoration: BoxDecoration(
        gradient: const LinearGradient(
          colors: [AppTheme.success, Color(0xFF20c997)],
          begin: Alignment.topLeft,
          end: Alignment.bottomRight,
        ),
        borderRadius: BorderRadius.circular(12),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            '📅 آج کی سمری',
            style: TextStyle(
                color: Colors.white,
                fontWeight: FontWeight.bold,
                fontSize: 16),
          ),
          const SizedBox(height: 10),
          _todayRow('🛒 فروخت', bp.todaySalesTotal),
          _todayRow('📦 خریداری', bp.todayPurchTotal, negative: true),
          _todayRow('📈 آج منافع', bp.todayGrossProfit),
          _todayRow('💸 اخراجات', bp.todayExpTotal, negative: true),
          _todayRow('🏷️ کسٹمر ڈسکاؤنٹ', 0.0,
              negative: true),
          _todayRow('🏷️ سپلائر ڈسکاؤنٹ', 0.0),
          _todayRow('🧾 کسٹمر ٹیکس', 0.0),
          if (adj['loss']! > 0)
            _todayRow('⚖️ وزن نقصان', adj['loss']!, negative: true),
          if (adj['gain']! > 0) _todayRow('⚖️ وزن فائدہ', adj['gain']!),
        ],
      ),
    );
  }

  Widget _todayRow(String label, double amount, {bool negative = false}) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 3),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label,
              style:
                  const TextStyle(color: Colors.white70, fontSize: 13)),
          Text(
            '${negative ? '-' : '+'}Rs. ${amount.toStringAsFixed(2)}',
            style: TextStyle(
              color: negative
                  ? Colors.red.shade200
                  : Colors.greenAccent.shade100,
              fontWeight: FontWeight.bold,
              fontSize: 13,
            ),
          ),
        ],
      ),
    );
  }

  // آج کی ٹرانزیکشنز
  Widget _buildTodayTransactions(BusinessProvider bp) {
    final todaySales = bp.todaySales;
    final todayPurchases = bp.todayPurchases;
    final todayExpenses = bp.todayExpList;
    final todayCash = bp.todayCashList;

    if (todaySales.isEmpty &&
        todayPurchases.isEmpty &&
        todayExpenses.isEmpty &&
        todayCash.isEmpty) {
      return Container(
        padding: const EdgeInsets.all(20),
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(12),
          border: Border.all(color: AppTheme.cardBorder),
        ),
        child: const Center(
          child: Text(
            '📭 آج کوئی ٹرانزیکشن نہیں',
            style: TextStyle(color: Colors.grey, fontSize: 14),
          ),
        ),
      );
    }

    return Container(
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: AppTheme.cardBorder),
      ),
      child: Column(
        children: [
          Container(
            padding: const EdgeInsets.all(12),
            decoration: const BoxDecoration(
              color: Color(0xFF17a2b8),
              borderRadius: BorderRadius.only(
                topLeft: Radius.circular(12),
                topRight: Radius.circular(12),
              ),
            ),
            child: const Row(
              children: [
                Icon(Icons.receipt_long, color: Colors.white),
                SizedBox(width: 8),
                Text(
                  '🕐 آج کی ٹرانزیکشنز',
                  style: TextStyle(
                      color: Colors.white,
                      fontWeight: FontWeight.bold,
                      fontSize: 15),
                ),
              ],
            ),
          ),

          // فروخت
          if (todaySales.isNotEmpty)
            _transactionSection(
              '🛒 فروخت (${todaySales.length})',
              todaySales.map((s) {
                final c = bp.getCust(s.customerId);
                final itemNames = s.items.map((i) {
                  final item = bp.getStock(i.itemId);
                  return '${item?.name ?? '?'} ${i.qty}${item?.unit ?? ''}';
                }).join(', ');
                return _txRow(
                  '${c?.name ?? '-'} | $itemNames',
                  'Rs. ${s.total.toStringAsFixed(2)}',
                  'نقد: ${s.cashReceived.toStringAsFixed(2)} | بقایا: ${s.creditAmount.toStringAsFixed(2)}',
                  AppTheme.success,
                );
              }).toList(),
            ),

          // خریداری
          if (todayPurchases.isNotEmpty)
            _transactionSection(
              '📦 خریداری (${todayPurchases.length})',
              todayPurchases.map((p) {
                final s = bp.getSupp(p.supplierId);
                final itemNames = p.items.map((i) {
                  final item = bp.getStock(i.itemId);
                  return '${item?.name ?? '?'} ${i.qty}${item?.unit ?? ''}';
                }).join(', ');
                return _txRow(
                  '${s?.name ?? '-'} | $itemNames',
                  'Rs. ${p.total.toStringAsFixed(2)}',
                  'نقد: ${p.cashPaid.toStringAsFixed(2)} | بقایا: ${p.creditAmount.toStringAsFixed(2)}',
                  AppTheme.warning,
                );
              }).toList(),
            ),

          // اخراجات
          if (todayExpenses.isNotEmpty)
            _transactionSection(
              '💸 اخراجات (${todayExpenses.length})',
              todayExpenses.map((e) => _txRow(
                    e.category,
                    'Rs. ${e.amount.toStringAsFixed(2)}',
                    e.note ?? '',
                    AppTheme.danger,
                  )).toList(),
            ),

          // کیش ان/آؤٹ
          if (todayCash.isNotEmpty)
            _transactionSection(
              '💰 کیش ان/آؤٹ (${todayCash.length})',
              todayCash.map((c) => _txRow(
                    c.type,
                    '${c.amount >= 0 ? '+' : ''}Rs. ${c.amount.toStringAsFixed(2)}',
                    c.note ?? '',
                    c.amount >= 0 ? AppTheme.success : AppTheme.danger,
                  )).toList(),
            ),
        ],
      ),
    );
  }

  Widget _transactionSection(String title, List<Widget> rows) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Container(
          padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
          color: AppTheme.lightBg,
          child: Text(
            title,
            style: const TextStyle(
                fontWeight: FontWeight.bold, fontSize: 13),
          ),
        ),
        ...rows,
      ],
    );
  }

  Widget _txRow(
      String title, String amount, String subtitle, Color color) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
      decoration: BoxDecoration(
        border: Border(
          bottom: BorderSide(color: AppTheme.cardBorder.withOpacity(0.4)),
          right: BorderSide(color: color, width: 3),
        ),
      ),
      child: Row(
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(title,
                    style: const TextStyle(
                        fontWeight: FontWeight.bold, fontSize: 13)),
                if (subtitle.isNotEmpty)
                  Text(subtitle,
                      style: TextStyle(
                          fontSize: 11, color: Colors.grey.shade500)),
              ],
            ),
          ),
          Text(
            amount,
            style: TextStyle(
                color: color,
                fontWeight: FontWeight.w900,
                fontSize: 13),
          ),
        ],
      ),
    );
  }

  // اسٹاک سمری
  Widget _buildStockSummary(BusinessProvider bp) {
    return Container(
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: AppTheme.cardBorder),
      ),
      child: Column(
        children: [
          InkWell(
            onTap: () => Navigator.pushNamed(context, '/stock'),
            child: Container(
              padding: const EdgeInsets.all(12),
              decoration: const BoxDecoration(
                color: AppTheme.info,
                borderRadius: BorderRadius.only(
                  topLeft: Radius.circular(12),
                  topRight: Radius.circular(12),
                ),
              ),
              child: Row(
                children: [
                  const Icon(Icons.inventory_2, color: Colors.white),
                  const SizedBox(width: 8),
                  const Text(
                    '📦 اسٹاک حالت',
                    style: TextStyle(
                        color: Colors.white,
                        fontWeight: FontWeight.bold,
                        fontSize: 15),
                  ),
                  const Spacer(),
                  Text(
                    '${bp.stockItems.length} آئٹمز',
                    style: const TextStyle(
                        color: Colors.white70, fontSize: 12),
                  ),
                ],
              ),
            ),
          ),
          ...bp.stockItems.take(5).map((item) {
            final profit = bp.calcItemProfit(item.id);
            return Container(
              padding:
                  const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
              decoration: BoxDecoration(
                color: item.isLowStock
                    ? AppTheme.danger.withOpacity(0.04)
                    : null,
                border: Border(
                  bottom: BorderSide(
                      color: AppTheme.cardBorder.withOpacity(0.4)),
                  right: BorderSide(
                      color: item.isLowStock
                          ? AppTheme.danger
                          : AppTheme.info,
                      width: 3),
                ),
              ),
              child: Row(
                children: [
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Row(
                          children: [
                            Text(
                              item.name,
                              style: const TextStyle(
                                  fontWeight: FontWeight.bold,
                                  fontSize: 14),
                            ),
                            if (item.isLowStock) ...[
                              const SizedBox(width: 6),
                              const StatusBadge(
                                  label: '⚠️ کم اسٹاک',
                                  color: AppTheme.danger),
                            ],
                          ],
                        ),
                        Text(
                          '${item.quantity} ${item.unit} | ریٹ: Rs. ${item.purchaseRate.toStringAsFixed(2)}',
                          style: TextStyle(
                              fontSize: 11,
                              color: Colors.grey.shade600),
                        ),
                      ],
                    ),
                  ),
                  Column(
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      Text(
                        'Rs. ${(item.quantity * item.purchaseRate).toStringAsFixed(0)}',
                        style: const TextStyle(
                            fontWeight: FontWeight.w900,
                            color: AppTheme.info,
                            fontSize: 13),
                      ),
                      Text(
                        'منافع: Rs. ${profit.toStringAsFixed(0)}',
                        style: const TextStyle(
                            fontSize: 10, color: AppTheme.success),
                      ),
                    ],
                  ),
                ],
              ),
            );
          }),
          if (bp.stockItems.length > 5)
            TextButton(
              onPressed: () =>
                  Navigator.pushNamed(context, '/stock'),
              child: Text(
                'مزید ${bp.stockItems.length - 5} آئٹمز دیکھیں →',
                style: const TextStyle(color: AppTheme.primary),
              ),
            ),
        ],
      ),
    );
  }
}



// ============================================================
// کسٹمرز اسکرین
// ============================================================

class CustomersScreen extends StatefulWidget {
  const CustomersScreen({super.key});

  @override
  State<CustomersScreen> createState() => _CustomersScreenState();
}

class _CustomersScreenState extends State<CustomersScreen> {
  final _searchCtrl = TextEditingController();
  String _searchQuery = '';

  @override
  void dispose() {
    _searchCtrl.dispose();
    super.dispose();
  }

  List<Customer> _filteredCustomers(List<Customer> all) {
    if (_searchQuery.isEmpty) return all;
    final q = _searchQuery.toLowerCase();
    return all.where((c) {
      return c.name.toLowerCase().contains(q) ||
          c.phone.contains(q) ||
          c.address.toLowerCase().contains(q);
    }).toList();
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(
      builder: (context, bp, _) {
        final customers = _filteredCustomers(bp.customers);
        final totalReceivables =
            bp.customers.fold(0.0, (s, c) => s + (c.balance > 0 ? c.balance : 0));
        final totalOverpaid =
            bp.customers.fold(0.0, (s, c) => s + (c.balance < 0 ? c.balance.abs() : 0));

        return Scaffold(
          appBar: AppBar(
            title: const Text('👥 کسٹمرز'),
            backgroundColor: AppTheme.primary,
            actions: [
              IconButton(
                icon: const Icon(Icons.add),
                onPressed: () => _showAddCustomerDialog(context, bp),
              ),
            ],
          ),
          body: Column(
            children: [
              // سمری
              Container(
                padding: const EdgeInsets.all(12),
                color: AppTheme.primary,
                child: Row(
                  children: [
                    Expanded(
                      child: _summaryChip('کل کسٹمرز',
                          '${bp.customers.length}', Icons.people, Colors.white),
                    ),
                    Expanded(
                      child: _summaryChip('واجبات',
                          'Rs. ${totalReceivables.toStringAsFixed(0)}',
                          Icons.arrow_downward, Colors.red.shade200),
                    ),
                    Expanded(
                      child: _summaryChip('اضافی',
                          'Rs. ${totalOverpaid.toStringAsFixed(0)}',
                          Icons.arrow_upward, Colors.green.shade200),
                    ),
                  ],
                ),
              ),

              // سرچ
              Padding(
                padding: const EdgeInsets.all(12),
                child: AppSearchBox(
                  controller: _searchCtrl,
                  hint: '🔍 کسٹمر تلاش کریں (نام، فون، پتہ)',
                  onChanged: (v) => setState(() => _searchQuery = v),
                ),
              ),

              // لسٹ
              Expanded(
                child: customers.isEmpty
                    ? EmptyState(
                        message: 'کوئی کسٹمر نہیں ملا',
                        icon: Icons.people_outline,
                        buttonLabel: '➕ کسٹمر شامل کریں',
                        onButtonPress: () =>
                            _showAddCustomerDialog(context, bp),
                      )
                    : ListView.builder(
                        padding: const EdgeInsets.symmetric(horizontal: 12),
                        itemCount: customers.length,
                        itemBuilder: (ctx, i) =>
                            _customerCard(ctx, customers[i], bp),
                      ),
              ),
            ],
          ),
          floatingActionButton: FloatingActionButton(
            onPressed: () => _showAddCustomerDialog(context, bp),
            backgroundColor: AppTheme.primary,
            child: const Icon(Icons.person_add, color: Colors.white),
          ),
        );
      },
    );
  }

  Widget _summaryChip(
      String label, String value, IconData icon, Color color) {
    return Column(
      children: [
        Icon(icon, color: color, size: 18),
        Text(value,
            style: TextStyle(
                color: color, fontWeight: FontWeight.bold, fontSize: 13)),
        Text(label,
            style: const TextStyle(color: Colors.white60, fontSize: 10)),
      ],
    );
  }

  Widget _customerCard(
      BuildContext context, Customer c, BusinessProvider bp) {
    final balanceColor = c.balance > 0
        ? AppTheme.danger
        : c.balance < 0
            ? AppTheme.success
            : Colors.grey;
    final balanceLabel = c.balance > 0
        ? 'واجب الادا'
        : c.balance < 0
            ? 'اضافی'
            : 'برابر';

    return Card(
      margin: const EdgeInsets.only(bottom: 10),
      child: InkWell(
        onTap: () => Navigator.push(
          context,
          MaterialPageRoute(
              builder: (_) => CustomerDetailScreen(customerId: c.id)),
        ),
        borderRadius: BorderRadius.circular(12),
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Row(
                children: [
                  CircleAvatar(
                    backgroundColor: AppTheme.primary.withOpacity(0.1),
                    child: Text(
                      c.name.isNotEmpty ? c.name[0] : '?',
                      style: const TextStyle(
                          color: AppTheme.primary,
                          fontWeight: FontWeight.bold),
                    ),
                  ),
                  const SizedBox(width: 10),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(c.name,
                            style: const TextStyle(
                                fontWeight: FontWeight.bold,
                                fontSize: 16)),
                        if (c.phone.isNotEmpty)
                          Text('📞 ${c.phone}',
                              style: TextStyle(
                                  fontSize: 12,
                                  color: Colors.grey.shade600)),
                        if (c.address.isNotEmpty)
                          Text('📍 ${c.address}',
                              style: TextStyle(
                                  fontSize: 12,
                                  color: Colors.grey.shade600)),
                      ],
                    ),
                  ),
                  Column(
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      Text(
                        'Rs. ${c.balance.abs().toStringAsFixed(2)}',
                        style: TextStyle(
                            color: balanceColor,
                            fontWeight: FontWeight.w900,
                            fontSize: 16),
                      ),
                      StatusBadge(label: balanceLabel, color: balanceColor),
                    ],
                  ),
                ],
              ),
              const SizedBox(height: 10),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: () =>
                          _showReceivePaymentDialog(context, bp, c),
                      icon: const Icon(Icons.payments, size: 16),
                      label: const Text('💵 وصولی',
                          style: TextStyle(fontSize: 12)),
                      style: OutlinedButton.styleFrom(
                          foregroundColor: AppTheme.success,
                          side: const BorderSide(color: AppTheme.success)),
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: () =>
                          _showDiscountDialog(context, bp, c),
                      icon: const Icon(Icons.discount, size: 16),
                      label: const Text('🏷️ ڈسکاؤنٹ',
                          style: TextStyle(fontSize: 12)),
                      style: OutlinedButton.styleFrom(
                          foregroundColor: AppTheme.warning,
                          side: const BorderSide(color: AppTheme.warning)),
                    ),
                  ),
                  const SizedBox(width: 8),
                  PopupMenuButton<String>(
                    onSelected: (v) {
                      if (v == 'tax')
                        _showTaxDialog(context, bp, c);
                      else if (v == 'edit')
                        _showEditCustomerDialog(context, bp, c);
                      else if (v == 'delete')
                        _deleteCustomer(context, bp, c);
                      else if (v == 'detail')
                        Navigator.push(
                          context,
                          MaterialPageRoute(
                              builder: (_) =>
                                  CustomerDetailScreen(customerId: c.id)),
                        );
                    },
                    itemBuilder: (_) => [
                      const PopupMenuItem(
                          value: 'detail',
                          child: Text('📋 تفصیل دیکھیں')),
                      const PopupMenuItem(
                          value: 'tax', child: Text('🧾 ٹیکس شامل')),
                      const PopupMenuItem(
                          value: 'edit', child: Text('✏️ ترمیم')),
                      const PopupMenuItem(
                          value: 'delete',
                          child: Text('🗑️ حذف',
                              style: TextStyle(color: AppTheme.danger))),
                    ],
                    child: Container(
                      padding: const EdgeInsets.all(8),
                      decoration: BoxDecoration(
                        border: Border.all(color: AppTheme.cardBorder),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: const Icon(Icons.more_vert, size: 18),
                    ),
                  ),
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }

  // کسٹمر شامل کریں
  void _showAddCustomerDialog(BuildContext context, BusinessProvider bp) {
    final nameCtrl = TextEditingController();
    final phoneCtrl = TextEditingController();
    final addressCtrl = TextEditingController();
    final formKey = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(
            bottom: MediaQuery.of(ctx).viewInsets.bottom,
            left: 20, right: 20, top: 20),
        child: Form(
          key: formKey,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              const Text('👤 نیا کسٹمر شامل کریں',
                  style: TextStyle(
                      fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 16),
              AppFormField(
                  label: '📛 نام',
                  controller: nameCtrl,
                  required: true),
              const SizedBox(height: 12),
              AppFormField(
                  label: '📞 فون',
                  controller: phoneCtrl,
                  keyboardType: TextInputType.phone),
              const SizedBox(height: 12),
              AppFormField(
                  label: '📍 پتہ',
                  controller: addressCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!formKey.currentState!.validate()) return;
                    await bp.addCustomer(Customer(
                      id: 0,
                      name: nameCtrl.text.trim(),
                      phone: phoneCtrl.text.trim(),
                      address: addressCtrl.text.trim(),
                      balance: 0,
                      transactions: [],
                      createdAt: DateTime.now().toIso8601String(),
                    ));
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted)
                      showSnackBar(context, '✅ کسٹمر شامل ہو گیا!');
                  },
                  child: const Text('💾 محفوظ کریں'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  // کسٹمر ترمیم
  void _showEditCustomerDialog(
      BuildContext context, BusinessProvider bp, Customer c) {
    final nameCtrl = TextEditingController(text: c.name);
    final phoneCtrl = TextEditingController(text: c.phone);
    final addressCtrl = TextEditingController(text: c.address);
    final formKey = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(
            bottom: MediaQuery.of(ctx).viewInsets.bottom,
            left: 20, right: 20, top: 20),
        child: Form(
          key: formKey,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('✏️ ${c.name} — ترمیم',
                  style: const TextStyle(
                      fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 16),
              AppFormField(
                  label: '📛 نام',
                  controller: nameCtrl,
                  required: true),
              const SizedBox(height: 12),
              AppFormField(
                  label: '📞 فون',
                  controller: phoneCtrl,
                  keyboardType: TextInputType.phone),
              const SizedBox(height: 12),
              AppFormField(label: '📍 پتہ', controller: addressCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!formKey.currentState!.validate()) return;
                    final updated = c.copyWith(
                      name: nameCtrl.text.trim(),
                      phone: phoneCtrl.text.trim(),
                      address: addressCtrl.text.trim(),
                    );
                    await bp.updateCustomer(updated);
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted)
                      showSnackBar(context, '✅ تبدیلیاں محفوظ!');
                  },
                  child: const Text('💾 محفوظ کریں'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  // وصولی ڈائیلاگ
  void _showReceivePaymentDialog(
      BuildContext context, BusinessProvider bp, Customer c) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final formKey = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(
            bottom: MediaQuery.of(ctx).viewInsets.bottom,
            left: 20, right: 20, top: 20),
        child: Form(
          key: formKey,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('💵 ${c.name} — وصولی',
                  style: const TextStyle(
                      fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 8),
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(
                  color: AppTheme.info.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(8),
                ),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Text('موجودہ بیلنس:',
                        style: TextStyle(fontWeight: FontWeight.bold)),
                    Text(
                        'Rs. ${c.balance.toStringAsFixed(2)}',
                        style: TextStyle(
                            color: c.balance > 0
                                ? AppTheme.danger
                                : AppTheme.success,
                            fontWeight: FontWeight.bold)),
                  ],
                ),
              ),
              const SizedBox(height: 12),
              AppFormField(
                label: '💰 رقم',
                controller: amountCtrl,
                keyboardType:
                    const TextInputType.numberWithOptions(decimal: true),
                required: true,
                validator: (v) {
                  if (v == null || v.isEmpty) return 'رقم ضروری ہے';
                  if ((double.tryParse(v) ?? 0) <= 0)
                    return 'درست رقم درج کریں';
                  return null;
                },
              ),
              const SizedBox(height: 12),
              DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
              const SizedBox(height: 12),
              AppFormField(label: '📝 نوٹ', controller: noteCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!formKey.currentState!.validate()) return;
                    final amount = double.parse(amountCtrl.text);
                    await bp.receiveCustomerPayment(
                        c.id, amount, dateCtrl.text, noteCtrl.text);
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted)
                      showSnackBar(context,
                          '✅ Rs. ${amount.toStringAsFixed(2)} وصول ہوا!');
                  },
                  style: ElevatedButton.styleFrom(
                      backgroundColor: AppTheme.success,
                      foregroundColor: Colors.white),
                  child: const Text('💾 محفوظ کریں'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  // ڈسکاؤنٹ ڈائیلاگ
  void _showDiscountDialog(
      BuildContext context, BusinessProvider bp, Customer c) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final formKey = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(
            bottom: MediaQuery.of(ctx).viewInsets.bottom,
            left: 20, right: 20, top: 20),
        child: Form(
          key: formKey,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('🏷️ ${c.name} — ڈسکاؤنٹ',
                  style: const TextStyle(
                      fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 8),
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(
                  color: AppTheme.warning.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(8),
                  border: Border.all(
                      color: AppTheme.warning.withOpacity(0.3)),
                ),
                child: const Text(
                  'ℹ️ ڈسکاؤنٹ دینے سے کسٹمر کا بیلنس کم ہو گا اور منافع میں کمی ہو گی۔',
                  style: TextStyle(fontSize: 12),
                ),
              ),
              const SizedBox(height: 12),
              AppFormField(
                label: '💰 ڈسکاؤنٹ رقم',
                controller: amountCtrl,
                keyboardType:
                    const TextInputType.numberWithOptions(decimal: true),
                required: true,
              ),
              const SizedBox(height: 12),
              DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
              const SizedBox(height: 12),
              AppFormField(
                  label: '📝 تفصیل', controller: noteCtrl, required: true),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!formKey.currentState!.validate()) return;
                    final amount = double.parse(amountCtrl.text);
                    await bp.addCustomerDiscount(
                        c.id, amount, dateCtrl.text, noteCtrl.text);
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted)
                      showSnackBar(context,
                          '✅ ڈسکاؤنٹ Rs. ${amount.toStringAsFixed(2)} شامل!');
                  },
                  style: ElevatedButton.styleFrom(
                      backgroundColor: AppTheme.warning,
                      foregroundColor: Colors.black),
                  child: const Text('💾 محفوظ کریں'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  // ٹیکس ڈائیلاگ
  void _showTaxDialog(
      BuildContext context, BusinessProvider bp, Customer c) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final formKey = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(
            bottom: MediaQuery.of(ctx).viewInsets.bottom,
            left: 20, right: 20, top: 20),
        child: Form(
          key: formKey,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('🧾 ${c.name} — ٹیکس شامل',
                  style: const TextStyle(
                      fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 8),
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(
                  color: AppTheme.info.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(8),
                  border: Border.all(color: AppTheme.info.withOpacity(0.3)),
                ),
                child: const Text(
                  'ℹ️ یہ رقم کسٹمر کے کھاتے میں جمع ہوگی (انہیں دینی ہوگی) اور منافع میں شامل ہوگی۔',
                  style: TextStyle(fontSize: 12),
                ),
              ),
              const SizedBox(height: 12),
              AppFormField(
                label: '💰 ٹیکس رقم',
                controller: amountCtrl,
                keyboardType:
                    const TextInputType.numberWithOptions(decimal: true),
                required: true,
              ),
              const SizedBox(height: 12),
              DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
              const SizedBox(height: 12),
              AppFormField(
                label: '📝 تفصیل',
                controller: noteCtrl,
                required: true,
                hint: 'مثلاً: اکٹروئی ٹیکس، ٹرانسپورٹ چارج',
              ),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!formKey.currentState!.validate()) return;
                    final amount = double.parse(amountCtrl.text);
                    await bp.addCustomerTax(
                        c.id, amount, dateCtrl.text, noteCtrl.text);
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted)
                      showSnackBar(context,
                          '✅ ٹیکس Rs. ${amount.toStringAsFixed(2)} شامل!');
                  },
                  style: ElevatedButton.styleFrom(
                      backgroundColor: AppTheme.info,
                      foregroundColor: Colors.white),
                  child: const Text('💾 محفوظ کریں'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> _deleteCustomer(
      BuildContext context, BusinessProvider bp, Customer c) async {
    final confirm = await showConfirmDialog(
      context,
      '🗑️ ${c.name} حذف کریں؟',
      'یہ کسٹمر اور اس کی تمام تاریخچہ ہمیشہ کے لیے حذف ہو جائے گی!',
    );
    if (confirm) {
      await bp.deleteCustomer(c.id);
      if (mounted) showSnackBar(context, '✅ کسٹمر حذف ہو گیا!');
    }
  }
}



// ============================================================
// کسٹمر تفصیل اسکرین — ٹرانزیکشن تاریخچہ
// ============================================================

class CustomerDetailScreen extends StatelessWidget {
  final int customerId;

  const CustomerDetailScreen({super.key, required this.customerId});

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(
      builder: (context, bp, _) {
        final Customer? c;
        try {
          c = bp.customers.firstWhere((x) => x.id == customerId, orElse: () => Customer(id:0,name:'حذف شدہ',phone:'',address:'',balance:0,transactions:[],createdAt:''));
        } catch (_) {
          return Scaffold(
            appBar: AppBar(title: const Text('کسٹمر نہیں ملا')),
            body: const Center(child: Text('کسٹمر حذف ہو چکا ہے۔')),
          );
        }

        // کسٹمر کی فروخت نکالیں
        final custSales =
            bp.sales.where((s) => s.customerId == customerId).toList();

        // بیلنس رنگ
        final balColor = c.balance > 0
            ? AppTheme.danger
            : c.balance < 0
                ? AppTheme.success
                : Colors.grey;
        final balLabel = c.balance > 0
            ? 'واجب الادا'
            : c.balance < 0
                ? 'اضافی ادائیگی'
                : 'برابر';

        // کل فروخت
        final totalSales =
            custSales.fold(0.0, (s, sale) => s + sale.total);
        final totalPaid =
            custSales.fold(0.0, (s, sale) => s + sale.cashReceived);

        return Scaffold(
          backgroundColor: AppTheme.lightBg,
          appBar: AppBar(
            title: Text(c.name),
            backgroundColor: AppTheme.primary,
            actions: [
              IconButton(
                icon: const Icon(Icons.print),
                tooltip: 'پرنٹ',
                onPressed: () =>
                    _printStatement(context, bp, c!, custSales),
              ),
            ],
          ),
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(14),
            child: Column(
              children: [
                // پروفائل کارڈ
                Card(
                  child: Padding(
                    padding: const EdgeInsets.all(16),
                    child: Column(
                      children: [
                        CircleAvatar(
                          radius: 34,
                          backgroundColor: AppTheme.primary.withOpacity(0.1),
                          child: Text(
                            c.name.isNotEmpty ? c.name[0] : '?',
                            style: const TextStyle(
                                fontSize: 28,
                                color: AppTheme.primary,
                                fontWeight: FontWeight.bold),
                          ),
                        ),
                        const SizedBox(height: 12),
                        Text(c.name,
                            style: const TextStyle(
                                fontSize: 20, fontWeight: FontWeight.bold)),
                        if (c.phone.isNotEmpty)
                          Text('📞 ${c.phone}',
                              style: const TextStyle(
                                  color: Colors.grey, fontSize: 13)),
                        if (c.address.isNotEmpty)
                          Text('📍 ${c.address}',
                              style: const TextStyle(
                                  color: Colors.grey, fontSize: 13)),
                        const SizedBox(height: 12),
                        // بیلنس
                        Container(
                          padding: const EdgeInsets.symmetric(
                              vertical: 12, horizontal: 20),
                          decoration: BoxDecoration(
                            color: balColor.withOpacity(0.1),
                            borderRadius: BorderRadius.circular(12),
                            border: Border.all(
                                color: balColor.withOpacity(0.3)),
                          ),
                          child: Column(
                            children: [
                              Text(
                                'Rs. ${c.balance.abs().toStringAsFixed(2)}',
                                style: TextStyle(
                                    color: balColor,
                                    fontSize: 28,
                                    fontWeight: FontWeight.w900),
                              ),
                              StatusBadge(label: balLabel, color: balColor),
                            ],
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
                const SizedBox(height: 12),

                // اسٹیٹس گرڈ
                Row(
                  children: [
                    Expanded(
                      child: _miniStatCard('کل فروخت',
                          'Rs. ${totalSales.toStringAsFixed(0)}',
                          AppTheme.primary, Icons.shopping_cart),
                    ),
                    const SizedBox(width: 10),
                    Expanded(
                      child: _miniStatCard('کل ادائیگی',
                          'Rs. ${totalPaid.toStringAsFixed(0)}',
                          AppTheme.success, Icons.payments),
                    ),
                    const SizedBox(width: 10),
                    Expanded(
                      child: _miniStatCard('فروخت تعداد',
                          '${custSales.length}',
                          AppTheme.info, Icons.receipt),
                    ),
                  ],
                ),
                const SizedBox(height: 12),

                // ٹرانزیکشن تاریخچہ
                Card(
                  child: Column(
                    children: [
                      Container(
                        padding: const EdgeInsets.all(12),
                        decoration: const BoxDecoration(
                          color: AppTheme.primary,
                          borderRadius: BorderRadius.only(
                              topLeft: Radius.circular(12),
                              topRight: Radius.circular(12)),
                        ),
                        child: const Row(
                          children: [
                            Icon(Icons.history, color: Colors.white),
                            SizedBox(width: 8),
                            Text('📋 لین دین تاریخچہ',
                                style: TextStyle(
                                    color: Colors.white,
                                    fontWeight: FontWeight.bold,
                                    fontSize: 15)),
                          ],
                        ),
                      ),
                      if (c.transactions.isEmpty)
                        const Padding(
                          padding: EdgeInsets.all(20),
                          child: Text('کوئی ٹرانزیکشن نہیں',
                              style: TextStyle(color: Colors.grey)),
                        )
                      else
                        ...c.transactions.reversed.map((t) {
                          final isPositive = t.amount < 0;
                          return TransactionTile(
                            date: t.date,
                            type: t.type,
                            amount: t.amount.abs(),
                            note: t.note,
                            isPositive: isPositive,
                          );
                        }),
                    ],
                  ),
                ),
                const SizedBox(height: 12),

                // فروخت تفصیل
                if (custSales.isNotEmpty) ...[
                  Card(
                    child: Column(
                      children: [
                        Container(
                          padding: const EdgeInsets.all(12),
                          decoration: const BoxDecoration(
                            color: AppTheme.success,
                            borderRadius: BorderRadius.only(
                                topLeft: Radius.circular(12),
                                topRight: Radius.circular(12)),
                          ),
                          child: const Row(
                            children: [
                              Icon(Icons.shopping_cart,
                                  color: Colors.white),
                              SizedBox(width: 8),
                              Text('🛒 فروخت تفصیل',
                                  style: TextStyle(
                                      color: Colors.white,
                                      fontWeight: FontWeight.bold,
                                      fontSize: 15)),
                            ],
                          ),
                        ),
                        ...custSales.take(10).map((sale) {
                          final itemDesc = sale.items.map((i) {
                            final item = bp.getStock(i.itemId);
                            return '${item?.name ?? '?'} ×${i.qty}${item?.unit ?? ''}@${i.rate.toStringAsFixed(0)}';
                          }).join(', ');
                          return Container(
                            padding: const EdgeInsets.symmetric(
                                horizontal: 12, vertical: 8),
                            decoration: BoxDecoration(
                              border: Border(
                                  bottom: BorderSide(
                                      color: AppTheme.cardBorder
                                          .withOpacity(0.4))),
                            ),
                            child: Row(
                              children: [
                                Expanded(
                                  child: Column(
                                    crossAxisAlignment:
                                        CrossAxisAlignment.start,
                                    children: [
                                      Text(sale.date,
                                          style: const TextStyle(
                                              fontSize: 11,
                                              color: Colors.grey)),
                                      Text(itemDesc,
                                          style: const TextStyle(
                                              fontSize: 12,
                                              fontWeight:
                                                  FontWeight.bold)),
                                      if (sale.note.isNotEmpty)
                                        Text(sale.note,
                                            style: const TextStyle(
                                                fontSize: 11,
                                                color: Colors.grey)),
                                    ],
                                  ),
                                ),
                                Column(
                                  crossAxisAlignment:
                                      CrossAxisAlignment.end,
                                  children: [
                                    Text(
                                      'Rs. ${sale.total.toStringAsFixed(2)}',
                                      style: const TextStyle(
                                          fontWeight: FontWeight.w900,
                                          color: AppTheme.primary,
                                          fontSize: 14),
                                    ),
                                    if (sale.creditAmount > 0)
                                      Text(
                                        'بقایا: ${sale.creditAmount.toStringAsFixed(2)}',
                                        style: const TextStyle(
                                            fontSize: 10,
                                            color: AppTheme.danger),
                                      ),
                                  ],
                                ),
                              ],
                            ),
                          );
                        }),
                      ],
                    ),
                  ),
                ],
                const SizedBox(height: 60),
              ],
            ),
          ),
        );
      },
    );
  }

  Widget _miniStatCard(
      String label, String value, Color color, IconData icon) {
    return Container(
      padding: const EdgeInsets.all(10),
      decoration: BoxDecoration(
        color: color.withOpacity(0.1),
        borderRadius: BorderRadius.circular(10),
        border: Border.all(color: color.withOpacity(0.3)),
      ),
      child: Column(
        children: [
          Icon(icon, color: color, size: 18),
          const SizedBox(height: 4),
          Text(value,
              style: TextStyle(
                  color: color,
                  fontWeight: FontWeight.bold,
                  fontSize: 12)),
          Text(label,
              style:
                  const TextStyle(color: Colors.grey, fontSize: 10)),
        ],
      ),
    );
  }

  void _printStatement(BuildContext context, BusinessProvider bp,
      Customer c, List<Sale> sales) {
    ScaffoldMessenger.of(context).showSnackBar(const SnackBar(
      content: Text('🖨️ پرنٹ فیچر جلد آ رہا ہے...'),
      backgroundColor: AppTheme.info,
    ));
  }
}



// ============================================================
// سپلائرز اسکرین
// ============================================================

class SuppliersScreen extends StatefulWidget {
  const SuppliersScreen({super.key});

  @override
  State<SuppliersScreen> createState() => _SuppliersScreenState();
}

class _SuppliersScreenState extends State<SuppliersScreen> {
  final _searchCtrl = TextEditingController();
  String _searchQuery = '';

  List<Supplier> _filtered(List<Supplier> all) {
    if (_searchQuery.isEmpty) return all;
    final q = _searchQuery.toLowerCase();
    return all.where((s) =>
        s.name.toLowerCase().contains(q) ||
        s.phone.contains(q) ||
        s.address.toLowerCase().contains(q)).toList();
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final suppliers = _filtered(bp.suppliers);
      final totalPayable = bp.suppliers
          .fold(0.0, (s, x) => s + (x.balance > 0 ? x.balance : 0));
      final totalOverpaid = bp.suppliers
          .fold(0.0, (s, x) => s + (x.balance < 0 ? x.balance.abs() : 0));

      return Scaffold(
        appBar: AppBar(
          title: const Text('🚚 سپلائرز'),
          backgroundColor: AppTheme.warning,
          foregroundColor: Colors.black,
          actions: [
            IconButton(
              icon: const Icon(Icons.add),
              onPressed: () => _showAddDialog(context, bp),
            ),
          ],
        ),
        body: Column(
          children: [
            // سمری
            Container(
              padding: const EdgeInsets.all(12),
              color: AppTheme.warning,
              child: Row(
                children: [
                  Expanded(child: _chip('سپلائرز', '${bp.suppliers.length}', Colors.black87)),
                  Expanded(child: _chip('واجبات', 'Rs. ${totalPayable.toStringAsFixed(0)}', AppTheme.danger)),
                  Expanded(child: _chip('اضافی', 'Rs. ${totalOverpaid.toStringAsFixed(0)}', AppTheme.success)),
                ],
              ),
            ),
            Padding(
              padding: const EdgeInsets.all(12),
              child: AppSearchBox(
                controller: _searchCtrl,
                hint: '🔍 سپلائر تلاش کریں',
                onChanged: (v) => setState(() => _searchQuery = v),
              ),
            ),
            Expanded(
              child: suppliers.isEmpty
                  ? EmptyState(
                      message: 'کوئی سپلائر نہیں ملا',
                      icon: Icons.local_shipping_outlined,
                      buttonLabel: '➕ سپلائر شامل کریں',
                      onButtonPress: () => _showAddDialog(context, bp),
                    )
                  : ListView.builder(
                      padding: const EdgeInsets.symmetric(horizontal: 12),
                      itemCount: suppliers.length,
                      itemBuilder: (ctx, i) => _card(ctx, suppliers[i], bp),
                    ),
            ),
          ],
        ),
        floatingActionButton: FloatingActionButton(
          onPressed: () => _showAddDialog(context, bp),
          backgroundColor: AppTheme.warning,
          foregroundColor: Colors.black,
          child: const Icon(Icons.add),
        ),
      );
    });
  }

  Widget _chip(String label, String value, Color color) {
    return Column(
      children: [
        Text(value, style: TextStyle(color: color, fontWeight: FontWeight.bold, fontSize: 13)),
        Text(label, style: const TextStyle(color: Colors.black54, fontSize: 10)),
      ],
    );
  }

  Widget _card(BuildContext context, Supplier s, BusinessProvider bp) {
    final balColor = s.balance > 0 ? AppTheme.danger : s.balance < 0 ? AppTheme.success : Colors.grey;
    final balLabel = s.balance > 0 ? 'واجب الادا' : s.balance < 0 ? 'اضافی' : 'برابر';

    return Card(
      margin: const EdgeInsets.only(bottom: 10),
      child: Padding(
        padding: const EdgeInsets.all(14),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                CircleAvatar(
                  backgroundColor: AppTheme.warning.withOpacity(0.2),
                  child: Text(s.name.isNotEmpty ? s.name[0] : '?',
                      style: const TextStyle(color: AppTheme.warning, fontWeight: FontWeight.bold)),
                ),
                const SizedBox(width: 10),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(s.name, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                      if (s.phone.isNotEmpty) Text('📞 ${s.phone}', style: TextStyle(fontSize: 12, color: Colors.grey.shade600)),
                      if (s.address.isNotEmpty) Text('📍 ${s.address}', style: TextStyle(fontSize: 12, color: Colors.grey.shade600)),
                    ],
                  ),
                ),
                Column(
                  crossAxisAlignment: CrossAxisAlignment.end,
                  children: [
                    Text('Rs. ${s.balance.abs().toStringAsFixed(2)}',
                        style: TextStyle(color: balColor, fontWeight: FontWeight.w900, fontSize: 16)),
                    StatusBadge(label: balLabel, color: balColor),
                  ],
                ),
              ],
            ),
            const SizedBox(height: 10),
            // ٹرانزیکشن تاریخچہ بٹن
            if (s.transactions.isNotEmpty) ...[
              ExpansionTile(
                title: Text('📋 تاریخچہ (${s.transactions.length})',
                    style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold)),
                tilePadding: EdgeInsets.zero,
                children: s.transactions.reversed.take(5).map((t) {
                  final isPositive = t.amount < 0;
                  return TransactionTile(
                    date: t.date,
                    type: t.type,
                    amount: t.amount.abs(),
                    note: t.note,
                    isPositive: isPositive,
                  );
                }).toList(),
              ),
            ],
            const SizedBox(height: 8),
            Row(
              children: [
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _showPayDialog(context, bp, s),
                    icon: const Icon(Icons.payment, size: 16),
                    label: const Text('💸 ادائیگی', style: TextStyle(fontSize: 12)),
                    style: OutlinedButton.styleFrom(foregroundColor: AppTheme.danger, side: const BorderSide(color: AppTheme.danger)),
                  ),
                ),
                const SizedBox(width: 8),
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _showDiscountDialog(context, bp, s),
                    icon: const Icon(Icons.discount, size: 16),
                    label: const Text('🏷️ ڈسکاؤنٹ', style: TextStyle(fontSize: 12)),
                    style: OutlinedButton.styleFrom(foregroundColor: AppTheme.success, side: const BorderSide(color: AppTheme.success)),
                  ),
                ),
                const SizedBox(width: 8),
                PopupMenuButton<String>(
                  onSelected: (v) {
                    if (v == 'edit') _showEditDialog(context, bp, s);
                    if (v == 'delete') _delete(context, bp, s);
                  },
                  itemBuilder: (_) => [
                    const PopupMenuItem(value: 'edit', child: Text('✏️ ترمیم')),
                    const PopupMenuItem(value: 'delete', child: Text('🗑️ حذف', style: TextStyle(color: AppTheme.danger))),
                  ],
                  child: Container(
                    padding: const EdgeInsets.all(8),
                    decoration: BoxDecoration(border: Border.all(color: AppTheme.cardBorder), borderRadius: BorderRadius.circular(8)),
                    child: const Icon(Icons.more_vert, size: 18),
                  ),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  void _showAddDialog(BuildContext context, BusinessProvider bp) {
    final nameCtrl = TextEditingController();
    final phoneCtrl = TextEditingController();
    final addressCtrl = TextEditingController();
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(
          key: key,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              const Text('🚚 نیا سپلائر', style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 16),
              AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
              const SizedBox(height: 12),
              AppFormField(label: '📞 فون', controller: phoneCtrl, keyboardType: TextInputType.phone),
              const SizedBox(height: 12),
              AppFormField(label: '📍 پتہ', controller: addressCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!key.currentState!.validate()) return;
                    await bp.addSupplier(Supplier(
                      id: 0, name: nameCtrl.text.trim(), phone: phoneCtrl.text.trim(),
                      address: addressCtrl.text.trim(), balance: 0,
                      transactions: [], createdAt: DateTime.now().toIso8601String(),
                    ));
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted) showSnackBar(context, '✅ سپلائر شامل!');
                  },
                  child: const Text('💾 محفوظ'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  void _showEditDialog(BuildContext context, BusinessProvider bp, Supplier s) {
    final nameCtrl = TextEditingController(text: s.name);
    final phoneCtrl = TextEditingController(text: s.phone);
    final addressCtrl = TextEditingController(text: s.address);
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(
          key: key,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('✏️ ${s.name}', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 16),
              AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
              const SizedBox(height: 12),
              AppFormField(label: '📞 فون', controller: phoneCtrl, keyboardType: TextInputType.phone),
              const SizedBox(height: 12),
              AppFormField(label: '📍 پتہ', controller: addressCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!key.currentState!.validate()) return;
                    await bp.updateSupplier(s.copyWith(
                      name: nameCtrl.text.trim(), phone: phoneCtrl.text.trim(), address: addressCtrl.text.trim(),
                    ));
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted) showSnackBar(context, '✅ تبدیلیاں محفوظ!');
                  },
                  child: const Text('💾 محفوظ'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  void _showPayDialog(BuildContext context, BusinessProvider bp, Supplier s) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(
          key: key,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('💸 ${s.name} — ادائیگی', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 8),
              Container(
                padding: const EdgeInsets.all(10),
                decoration: BoxDecoration(color: AppTheme.info.withOpacity(0.1), borderRadius: BorderRadius.circular(8)),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    const Text('موجودہ بیلنس:', style: TextStyle(fontWeight: FontWeight.bold)),
                    Text('Rs. ${s.balance.toStringAsFixed(2)}', style: TextStyle(color: s.balance > 0 ? AppTheme.danger : AppTheme.success, fontWeight: FontWeight.bold)),
                  ],
                ),
              ),
              const SizedBox(height: 12),
              AppFormField(
                label: '💰 رقم', controller: amountCtrl, required: true,
                keyboardType: const TextInputType.numberWithOptions(decimal: true),
                validator: (v) => (v == null || v.isEmpty || (double.tryParse(v) ?? 0) <= 0) ? 'درست رقم درج کریں' : null,
              ),
              const SizedBox(height: 12),
              DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
              const SizedBox(height: 12),
              AppFormField(label: '📝 نوٹ', controller: noteCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!key.currentState!.validate()) return;
                    final amount = double.parse(amountCtrl.text);
                    await bp.paySupplier(s.id, amount, dateCtrl.text, noteCtrl.text);
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted) showSnackBar(context, '✅ Rs. ${amount.toStringAsFixed(2)} ادا ہوا!');
                  },
                  style: ElevatedButton.styleFrom(backgroundColor: AppTheme.danger, foregroundColor: Colors.white),
                  child: const Text('💾 محفوظ'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  void _showDiscountDialog(BuildContext context, BusinessProvider bp, Supplier s) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(
          key: key,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text('🏷️ ${s.name} — ڈسکاؤنٹ', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 12),
              AppFormField(
                label: '💰 ڈسکاؤنٹ رقم', controller: amountCtrl, required: true,
                keyboardType: const TextInputType.numberWithOptions(decimal: true),
              ),
              const SizedBox(height: 12),
              DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
              const SizedBox(height: 12),
              AppFormField(label: '📝 تفصیل', controller: noteCtrl, required: true),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!key.currentState!.validate()) return;
                    final amount = double.parse(amountCtrl.text);
                    await bp.addSupplierDiscount(s.id, amount, dateCtrl.text, noteCtrl.text);
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted) showSnackBar(context, '✅ ڈسکاؤنٹ شامل!');
                  },
                  style: ElevatedButton.styleFrom(backgroundColor: AppTheme.success, foregroundColor: Colors.white),
                  child: const Text('💾 محفوظ'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> _delete(BuildContext context, BusinessProvider bp, Supplier s) async {
    final ok = await showConfirmDialog(context, '🗑️ ${s.name} حذف کریں؟',
        'اس سپلائر اور تمام تاریخچہ کو ہمیشہ کے لیے حذف کریں؟');
    if (ok) {
      await bp.deleteSupplier(s.id);
      if (mounted) showSnackBar(context, '✅ سپلائر حذف!');
    }
  }
}



// ============================================================
// ادھار مینجمنٹ اسکرین
// balance > 0 = ہم نے دینا ہے (Given)
// balance < 0 = ہم نے لینا ہے (Taken)
// ============================================================

class UdharScreen extends StatefulWidget {
  const UdharScreen({super.key});
  @override
  State<UdharScreen> createState() => _UdharScreenState();
}

class _UdharScreenState extends State<UdharScreen>
    with SingleTickerProviderStateMixin {
  late TabController _tabCtrl;
  final _searchCtrl = TextEditingController();
  String _searchQuery = '';

  @override
  void initState() {
    super.initState();
    _tabCtrl = TabController(length: 3, vsync: this);
  }

  @override
  void dispose() {
    _tabCtrl.dispose();
    _searchCtrl.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final allPersons = bp.udharPersons;
      final given = allPersons.where((u) => u.balance > 0).toList();
      final taken = allPersons.where((u) => u.balance < 0).toList();
      final settled = allPersons.where((u) => u.balance == 0).toList();

      final totalGiven = given.fold(0.0, (s, u) => s + u.balance);
      final totalTaken = taken.fold(0.0, (s, u) => s + u.balance.abs());

      final q = _searchQuery.toLowerCase();
      List<UdharPerson> current;
      if (_tabCtrl.index == 0) {
        current = _searchQuery.isEmpty
            ? given
            : given.where((u) => u.name.toLowerCase().contains(q)).toList();
      } else if (_tabCtrl.index == 1) {
        current = _searchQuery.isEmpty
            ? taken
            : taken.where((u) => u.name.toLowerCase().contains(q)).toList();
      } else {
        current = _searchQuery.isEmpty
            ? settled
            : settled.where((u) => u.name.toLowerCase().contains(q)).toList();
      }

      return Scaffold(
        appBar: AppBar(
          title: const Text('🤝 ادھار مینجمنٹ'),
          backgroundColor: AppTheme.purple,
          actions: [
            IconButton(
              icon: const Icon(Icons.add),
              onPressed: () => _showAddDialog(context, bp),
            ),
          ],
          bottom: TabBar(
            controller: _tabCtrl,
            onTap: (_) => setState(() {}),
            indicatorColor: Colors.white,
            labelColor: Colors.white,
            unselectedLabelColor: Colors.white60,
            tabs: [
              Tab(text: '📤 دیا (${given.length})'),
              Tab(text: '📥 لیا (${taken.length})'),
              Tab(text: '✅ برابر (${settled.length})'),
            ],
          ),
        ),
        body: Column(
          children: [
            // سمری
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
              color: AppTheme.purple,
              child: Row(
                children: [
                  Expanded(child: _chip('کل دیا', 'Rs. ${totalGiven.toStringAsFixed(0)}', Colors.orange.shade200)),
                  Expanded(child: _chip('کل لیا', 'Rs. ${totalTaken.toStringAsFixed(0)}', Colors.green.shade200)),
                  Expanded(child: _chip('خالص', 'Rs. ${(totalGiven - totalTaken).abs().toStringAsFixed(0)}',
                      totalGiven >= totalTaken ? Colors.orange.shade200 : Colors.green.shade200)),
                ],
              ),
            ),
            Padding(
              padding: const EdgeInsets.all(12),
              child: AppSearchBox(
                controller: _searchCtrl,
                hint: '🔍 نام سے تلاش',
                onChanged: (v) => setState(() => _searchQuery = v),
              ),
            ),
            Expanded(
              child: current.isEmpty
                  ? EmptyState(
                      message: 'کوئی ریکارڈ نہیں',
                      icon: Icons.handshake_outlined,
                      buttonLabel: '➕ شامل کریں',
                      onButtonPress: () => _showAddDialog(context, bp),
                    )
                  : ListView.builder(
                      padding: const EdgeInsets.symmetric(horizontal: 12),
                      itemCount: current.length,
                      itemBuilder: (ctx, i) => _card(ctx, current[i], bp),
                    ),
            ),
          ],
        ),
        floatingActionButton: FloatingActionButton(
          onPressed: () => _showAddDialog(context, bp),
          backgroundColor: AppTheme.purple,
          child: const Icon(Icons.person_add, color: Colors.white),
        ),
      );
    });
  }

  Widget _chip(String label, String value, Color color) => Column(
    children: [
      Text(value, style: TextStyle(color: color, fontWeight: FontWeight.bold, fontSize: 13)),
      Text(label, style: const TextStyle(color: Colors.white60, fontSize: 10)),
    ],
  );

  Widget _card(BuildContext context, UdharPerson u, BusinessProvider bp) {
    final isGiven = u.balance > 0;
    final isSettled = u.balance == 0;
    final color = isSettled ? Colors.grey : isGiven ? AppTheme.orange : AppTheme.success;
    final label = isSettled ? 'برابر' : isGiven ? 'دیا (واپس لینا ہے)' : 'لیا (واپس دینا ہے)';

    return Card(
      margin: const EdgeInsets.only(bottom: 10),
      child: Padding(
        padding: const EdgeInsets.all(14),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                CircleAvatar(
                  backgroundColor: color.withOpacity(0.1),
                  child: Text(u.name.isNotEmpty ? u.name[0] : '?',
                      style: TextStyle(color: color, fontWeight: FontWeight.bold)),
                ),
                const SizedBox(width: 10),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(u.name, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                      if (u.phone.isNotEmpty)
                        Text('📞 ${u.phone}', style: TextStyle(fontSize: 12, color: Colors.grey.shade600)),
                    ],
                  ),
                ),
                Column(
                  crossAxisAlignment: CrossAxisAlignment.end,
                  children: [
                    Text('Rs. ${u.balance.abs().toStringAsFixed(2)}',
                        style: TextStyle(color: color, fontWeight: FontWeight.w900, fontSize: 16)),
                    StatusBadge(label: label, color: color),
                  ],
                ),
              ],
            ),
            if (u.transactions.isNotEmpty) ...[
              const SizedBox(height: 8),
              ExpansionTile(
                title: Text('📋 تاریخچہ (${u.transactions.length})',
                    style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold)),
                tilePadding: EdgeInsets.zero,
                children: u.transactions.reversed.take(5).map((t) {
                  final pos = t.amount < 0 || t.type == 'واپسی';
                  return TransactionTile(
                    date: t.date, type: t.type, amount: t.amount.abs(),
                    note: t.note, isPositive: pos,
                  );
                }).toList(),
              ),
            ],
            const SizedBox(height: 8),
            Wrap(
              spacing: 8,
              runSpacing: 6,
              children: [
                if (!isGiven || isSettled)
                  _actionBtn('📤 ادھار دیں', AppTheme.orange, () => _giveUdhar(context, bp, u)),
                if (isGiven || isSettled)
                  _actionBtn('📥 ادھار لیں', AppTheme.success, () => _takeUdhar(context, bp, u)),
                _actionBtn('↩️ واپسی', AppTheme.primary, () => _returnUdhar(context, bp, u)),
                _actionBtn('✏️ ترمیم', AppTheme.info, () => _showEditDialog(context, bp, u)),
                _actionBtn('🗑️ حذف', AppTheme.danger, () => _delete(context, bp, u)),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Widget _actionBtn(String label, Color color, VoidCallback onTap) =>
      OutlinedButton(
        onPressed: onTap,
        style: OutlinedButton.styleFrom(
          foregroundColor: color,
          side: BorderSide(color: color),
          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
          textStyle: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold),
        ),
        child: Text(label),
      );

  void _showAddDialog(BuildContext context, BusinessProvider bp) {
    final nameCtrl = TextEditingController();
    final phoneCtrl = TextEditingController();
    final addressCtrl = TextEditingController();
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(
          key: key,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              const Text('🤝 نیا شخص شامل کریں', style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
              const SizedBox(height: 16),
              AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
              const SizedBox(height: 12),
              AppFormField(label: '📞 فون', controller: phoneCtrl, keyboardType: TextInputType.phone),
              const SizedBox(height: 12),
              AppFormField(label: '📍 پتہ', controller: addressCtrl),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!key.currentState!.validate()) return;
                    await bp.addUdharPerson(UdharPerson(
                      id: 0, name: nameCtrl.text.trim(), phone: phoneCtrl.text.trim(),
                      address: addressCtrl.text.trim(), balance: 0,
                      transactions: [], createdAt: DateTime.now().toIso8601String(),
                    ));
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted) showSnackBar(context, '✅ شخص شامل!');
                  },
                  child: const Text('💾 محفوظ'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  void _showEditDialog(BuildContext context, BusinessProvider bp, UdharPerson u) {
    final nameCtrl = TextEditingController(text: u.name);
    final phoneCtrl = TextEditingController(text: u.phone);
    final addressCtrl = TextEditingController(text: u.address);
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key,
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            Text('✏️ ${u.name}', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
            const SizedBox(height: 16),
            AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
            const SizedBox(height: 12),
            AppFormField(label: '📞 فون', controller: phoneCtrl, keyboardType: TextInputType.phone),
            const SizedBox(height: 12),
            AppFormField(label: '📍 پتہ', controller: addressCtrl),
            const SizedBox(height: 20),
            SizedBox(width: double.infinity, child: ElevatedButton(
              onPressed: () async {
                if (!key.currentState!.validate()) return;
                await bp.updateUdharPerson(u.copyWith(name: nameCtrl.text.trim(), phone: phoneCtrl.text.trim(), address: addressCtrl.text.trim()));
                if (ctx.mounted) Navigator.pop(ctx);
                if (mounted) showSnackBar(context, '✅ محفوظ!');
              },
              child: const Text('💾 محفوظ'),
            )),
            const SizedBox(height: 16),
          ]),
        ),
      ),
    );
  }

  void _giveUdhar(BuildContext context, BusinessProvider bp, UdharPerson u) {
    _udharDialog(context, bp, u, '📤 ادھار دیں (کیش آؤٹ)',
        AppTheme.orange, (amount, date, note) async {
      await bp.giveUdhar(u.id, amount, date, note);
      if (mounted) showSnackBar(context, '✅ Rs. ${amount.toStringAsFixed(2)} ادھار دیا!');
    });
  }

  void _takeUdhar(BuildContext context, BusinessProvider bp, UdharPerson u) {
    _udharDialog(context, bp, u, '📥 ادھار لیں (کیش ان)',
        AppTheme.success, (amount, date, note) async {
      await bp.takeUdhar(u.id, amount, date, note);
      if (mounted) showSnackBar(context, '✅ Rs. ${amount.toStringAsFixed(2)} ادھار لیا!');
    });
  }

  void _returnUdhar(BuildContext context, BusinessProvider bp, UdharPerson u) {
    final isGiven = u.balance > 0;
    _udharDialog(context, bp, u,
        isGiven ? '↩️ واپسی (آپ نے دیا تھا، واپس آیا)' : '↩️ واپسی (آپ نے لیا تھا، واپس دیں)',
        AppTheme.primary, (amount, date, note) async {
      await bp.returnUdhar(u.id, amount, isGiven, date, note);
      if (mounted) showSnackBar(context, '✅ واپسی Rs. ${amount.toStringAsFixed(2)} محفوظ!');
    });
  }

  void _udharDialog(BuildContext context, BusinessProvider bp, UdharPerson u,
      String title, Color color,
      Future<void> Function(double, String, String) onSave) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key,
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            Text(title, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold), textAlign: TextAlign.center),
            const SizedBox(height: 8),
            Container(
              padding: const EdgeInsets.all(10),
              decoration: BoxDecoration(color: color.withOpacity(0.1), borderRadius: BorderRadius.circular(8)),
              child: Row(mainAxisAlignment: MainAxisAlignment.spaceBetween, children: [
                const Text('موجودہ بیلنس:', style: TextStyle(fontWeight: FontWeight.bold)),
                Text('Rs. ${u.balance.abs().toStringAsFixed(2)}', style: TextStyle(color: color, fontWeight: FontWeight.bold)),
              ]),
            ),
            const SizedBox(height: 12),
            AppFormField(
              label: '💰 رقم', controller: amountCtrl, required: true,
              keyboardType: const TextInputType.numberWithOptions(decimal: true),
              validator: (v) => (v == null || v.isEmpty || (double.tryParse(v) ?? 0) <= 0) ? 'درست رقم درج کریں' : null,
            ),
            const SizedBox(height: 12),
            DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
            const SizedBox(height: 12),
            AppFormField(label: '📝 نوٹ', controller: noteCtrl),
            const SizedBox(height: 20),
            SizedBox(
              width: double.infinity,
              child: ElevatedButton(
                onPressed: () async {
                  if (!key.currentState!.validate()) return;
                  await onSave(double.parse(amountCtrl.text), dateCtrl.text, noteCtrl.text);
                  if (ctx.mounted) Navigator.pop(ctx);
                },
                style: ElevatedButton.styleFrom(backgroundColor: color, foregroundColor: Colors.white),
                child: const Text('💾 محفوظ'),
              ),
            ),
            const SizedBox(height: 16),
          ]),
        ),
      ),
    );
  }

  Future<void> _delete(BuildContext context, BusinessProvider bp, UdharPerson u) async {
    final ok = await showConfirmDialog(context, '🗑️ ${u.name} حذف کریں؟', 'تمام تاریخچہ حذف ہو جائے گا!');
    if (ok) {
      await bp.deleteUdharPerson(u.id);
      if (mounted) showSnackBar(context, '✅ حذف!');
    }
  }
}



// ============================================================
// اسٹاک مینجمنٹ اسکرین
// ============================================================

class StockScreen extends StatefulWidget {
  const StockScreen({super.key});
  @override
  State<StockScreen> createState() => _StockScreenState();
}

class _StockScreenState extends State<StockScreen>
    with SingleTickerProviderStateMixin {
  late TabController _tabs;
  final _searchCtrl = TextEditingController();
  String _search = '';
  String _sortBy = 'name';

  @override
  void initState() {
    super.initState();
    _tabs = TabController(length: 3, vsync: this);
  }

  @override
  void dispose() {
    _tabs.dispose();
    _searchCtrl.dispose();
    super.dispose();
  }

  List<StockItem> _applyFilters(List<StockItem> all, {bool lowOnly = false}) {
    var items = all.where((i) {
      if (_search.isNotEmpty) {
        return i.name.toLowerCase().contains(_search.toLowerCase()) ||
            i.category.toLowerCase().contains(_search.toLowerCase());
      }
      return true;
    }).toList();
    if (lowOnly) items = items.where((i) => i.isLowStock).toList();
    items.sort((a, b) {
      switch (_sortBy) {
        case 'qty':  return b.quantity.compareTo(a.quantity);
        case 'value': return (b.quantity * b.purchaseRate).compareTo(a.quantity * a.purchaseRate);
        default:     return a.name.compareTo(b.name);
      }
    });
    return items;
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final allItems = bp.stockItems;
      final lowItems = allItems.where((i) => i.isLowStock).toList();

      return Scaffold(
        appBar: AppBar(
          title: const Text('📦 اسٹاک مینجمنٹ'),
          backgroundColor: AppTheme.info,
          actions: [
            PopupMenuButton<String>(
              icon: const Icon(Icons.sort, color: Colors.white),
              onSelected: (v) => setState(() => _sortBy = v),
              itemBuilder: (_) => [
                const PopupMenuItem(value: 'name', child: Text('نام کے مطابق')),
                const PopupMenuItem(value: 'qty', child: Text('مقدار کے مطابق')),
                const PopupMenuItem(value: 'value', child: Text('مالیت کے مطابق')),
              ],
            ),
            IconButton(icon: const Icon(Icons.add), onPressed: () => _showAddStockDialog(context, bp)),
          ],
          bottom: TabBar(
            controller: _tabs,
            onTap: (_) => setState(() {}),
            indicatorColor: Colors.white,
            labelColor: Colors.white,
            unselectedLabelColor: Colors.white60,
            tabs: [
              Tab(text: 'تمام (${allItems.length})'),
              Tab(text: '⚠️ کم (${lowItems.length})'),
              Tab(text: '⚖️ ایڈجسٹمنٹ'),
            ],
          ),
        ),
        body: Column(
          children: [
            // سمری
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
              color: AppTheme.info,
              child: Row(
                children: [
                  Expanded(child: _chip('کل آئٹمز', '${allItems.length}', Colors.white)),
                  Expanded(child: _chip('کم اسٹاک', '${lowItems.length}', Colors.orange.shade200)),
                  Expanded(child: _chip('کل مالیت', 'Rs. ${bp.totalStockValue.toStringAsFixed(0)}', Colors.green.shade200)),
                ],
              ),
            ),
            if (_tabs.index != 2) ...[
              Padding(
                padding: const EdgeInsets.all(12),
                child: AppSearchBox(
                  controller: _searchCtrl,
                  hint: '🔍 آئٹم یا کیٹیگری تلاش کریں',
                  onChanged: (v) => setState(() => _search = v),
                ),
              ),
            ],
            Expanded(
              child: TabBarView(
                controller: _tabs,
                children: [
                  _buildItemList(context, bp, _applyFilters(allItems)),
                  _buildItemList(context, bp, _applyFilters(allItems, lowOnly: true)),
                  _buildAdjustmentTab(context, bp),
                ],
              ),
            ),
          ],
        ),
        floatingActionButton: _tabs.index != 2
            ? FloatingActionButton(
                onPressed: () => _showAddStockDialog(context, bp),
                backgroundColor: AppTheme.info,
                child: const Icon(Icons.add, color: Colors.white),
              )
            : null,
      );
    });
  }

  Widget _chip(String label, String value, Color color) => Column(children: [
    Text(value, style: TextStyle(color: color, fontWeight: FontWeight.bold, fontSize: 13)),
    Text(label, style: const TextStyle(color: Colors.white60, fontSize: 10)),
  ]);

  Widget _buildItemList(BuildContext context, BusinessProvider bp, List<StockItem> items) {
    if (items.isEmpty) {
      return EmptyState(
        message: 'کوئی آئٹم نہیں ملا',
        icon: Icons.inventory_2_outlined,
        buttonLabel: '➕ آئٹم شامل کریں',
        onButtonPress: () => _showAddStockDialog(context, bp),
      );
    }
    return ListView.builder(
      padding: const EdgeInsets.symmetric(horizontal: 12),
      itemCount: items.length,
      itemBuilder: (ctx, i) => _stockCard(ctx, bp, items[i]),
    );
  }

  Widget _stockCard(BuildContext context, BusinessProvider bp, StockItem item) {
    final profit = bp.calcItemProfit(item.id);
    final value = item.quantity * item.purchaseRate;
    final lastSupplier = item.lastSupplierId != null ? bp.getSupp(item.lastSupplierId!) : null;

    return Card(
      margin: const EdgeInsets.only(bottom: 10),
      child: Padding(
        padding: const EdgeInsets.all(14),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Container(
                  padding: const EdgeInsets.all(10),
                  decoration: BoxDecoration(
                    color: item.isLowStock ? AppTheme.danger.withOpacity(0.1) : AppTheme.info.withOpacity(0.1),
                    borderRadius: BorderRadius.circular(10),
                  ),
                  child: Icon(Icons.inventory_2, color: item.isLowStock ? AppTheme.danger : AppTheme.info),
                ),
                const SizedBox(width: 10),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Text(item.name, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                          const SizedBox(width: 6),
                          if (item.isLowStock) const StatusBadge(label: '⚠️ کم', color: AppTheme.danger),
                        ],
                      ),
                      if (item.category.isNotEmpty)
                        Text('📂 ${item.category}', style: TextStyle(fontSize: 12, color: Colors.grey.shade600)),
                      if (lastSupplier != null)
                        Text('🚚 ${lastSupplier.name}', style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                    ],
                  ),
                ),
                Column(
                  crossAxisAlignment: CrossAxisAlignment.end,
                  children: [
                    Text('${item.quantity} ${item.unit}',
                        style: TextStyle(
                            fontWeight: FontWeight.w900, fontSize: 18,
                            color: item.isLowStock ? AppTheme.danger : AppTheme.info)),
                    Text('Rs. ${value.toStringAsFixed(0)}',
                        style: const TextStyle(fontSize: 12, color: Colors.grey)),
                  ],
                ),
              ],
            ),
            const SizedBox(height: 10),

            // اسٹیٹس بار
            Row(
              children: [
                _infoTag('ریٹ', 'Rs. ${item.purchaseRate.toStringAsFixed(2)}', AppTheme.primary),
                const SizedBox(width: 8),
                _infoTag('الرٹ', '${item.alertLimit} ${item.unit}', AppTheme.warning),
                const SizedBox(width: 8),
                _infoTag('منافع', 'Rs. ${profit.toStringAsFixed(0)}', AppTheme.success),
              ],
            ),
            const SizedBox(height: 10),

            // خریداری تاریخچہ
            if (item.purchaseHistory.isNotEmpty) ...[
              ExpansionTile(
                title: Text('📜 خریداری تاریخچہ (${item.purchaseHistory.length})',
                    style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)),
                tilePadding: EdgeInsets.zero,
                childrenPadding: EdgeInsets.zero,
                children: item.purchaseHistory.reversed.take(5).map((ph) {
                  final supplier = bp.getSupp(ph.supplierId);
                  return Container(
                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 6),
                    decoration: BoxDecoration(
                        border: Border(bottom: BorderSide(color: AppTheme.cardBorder.withOpacity(0.3)))),
                    child: Row(
                      children: [
                        Expanded(child: Text('${ph.date} — ${supplier?.name ?? '?'}',
                            style: const TextStyle(fontSize: 11))),
                        Text('${ph.quantity} ${item.unit} @ Rs. ${ph.rate.toStringAsFixed(2)}',
                            style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold)),
                      ],
                    ),
                  );
                }).toList(),
              ),
            ],

            const SizedBox(height: 8),
            Row(
              children: [
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _showAdjDialog(context, bp, item),
                    icon: const Icon(Icons.tune, size: 14),
                    label: const Text('⚖️ ایڈجسٹ', style: TextStyle(fontSize: 12)),
                    style: OutlinedButton.styleFrom(foregroundColor: AppTheme.purple, side: const BorderSide(color: AppTheme.purple)),
                  ),
                ),
                const SizedBox(width: 8),
                IconButton(
                  icon: const Icon(Icons.edit, color: AppTheme.orange),
                  onPressed: () => _showEditDialog(context, bp, item),
                ),
                IconButton(
                  icon: const Icon(Icons.delete, color: AppTheme.danger),
                  onPressed: () => _deleteItem(context, bp, item),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  Widget _infoTag(String label, String value, Color color) => Container(
    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
    decoration: BoxDecoration(color: color.withOpacity(0.1), borderRadius: BorderRadius.circular(6)),
    child: Column(
      children: [
        Text(label, style: TextStyle(fontSize: 9, color: color)),
        Text(value, style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: color)),
      ],
    ),
  );

  // وزن ایڈجسٹمنٹ ٹیب
  Widget _buildAdjustmentTab(BuildContext context, BusinessProvider bp) {
    final adjs = bp.stockAdj;
    final totalLoss = adjs.where((a) => a.adjType == 'loss').fold(0.0, (s, a) => s + a.amount);
    final totalGain = adjs.where((a) => a.adjType == 'gain').fold(0.0, (s, a) => s + a.amount);

    return Column(
      children: [
        // سمری
        Container(
          margin: const EdgeInsets.all(12),
          padding: const EdgeInsets.all(12),
          decoration: BoxDecoration(
            gradient: const LinearGradient(colors: [AppTheme.purple, Color(0xFF9c27b0)]),
            borderRadius: BorderRadius.circular(12),
          ),
          child: Row(
            children: [
              Expanded(child: Column(children: [
                const Text('کل نقصان', style: TextStyle(color: Colors.white70, fontSize: 12)),
                Text('Rs. ${totalLoss.toStringAsFixed(0)}', style: const TextStyle(color: Colors.orangeAccent, fontWeight: FontWeight.bold)),
              ])),
              Expanded(child: Column(children: [
                const Text('کل فائدہ', style: TextStyle(color: Colors.white70, fontSize: 12)),
                Text('Rs. ${totalGain.toStringAsFixed(0)}', style: const TextStyle(color: Colors.greenAccent, fontWeight: FontWeight.bold)),
              ])),
              ElevatedButton(
                onPressed: () => _showNewAdjDialog(context, bp),
                style: ElevatedButton.styleFrom(backgroundColor: Colors.white, foregroundColor: AppTheme.purple),
                child: const Text('➕ نئی'),
              ),
            ],
          ),
        ),
        Expanded(
          child: adjs.isEmpty
              ? const EmptyState(message: 'کوئی ایڈجسٹمنٹ نہیں', icon: Icons.tune_outlined)
              : ListView.builder(
                  padding: const EdgeInsets.symmetric(horizontal: 12),
                  itemCount: adjs.length,
                  itemBuilder: (ctx, i) {
                    final a = adjs[i];
                    final isLoss = a.adjType == 'loss';
                    return Card(
                      margin: const EdgeInsets.only(bottom: 8),
                      child: ListTile(
                        leading: CircleAvatar(
                          backgroundColor: (isLoss ? AppTheme.danger : AppTheme.success).withOpacity(0.1),
                          child: Icon(isLoss ? Icons.trending_down : Icons.trending_up,
                              color: isLoss ? AppTheme.danger : AppTheme.success),
                        ),
                        title: Text('${a.itemName} — ${isLoss ? 'نقصان' : 'فائدہ'}',
                            style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                        subtitle: Text('${a.qty} ${a.unit} × Rs. ${a.rate.toStringAsFixed(2)}\n${a.date}${a.note.isNotEmpty ? ' | ${a.note}' : ''}',
                            style: const TextStyle(fontSize: 11)),
                        trailing: Text('Rs. ${a.amount.toStringAsFixed(0)}',
                            style: TextStyle(color: isLoss ? AppTheme.danger : AppTheme.success, fontWeight: FontWeight.bold)),
                      ),
                    );
                  },
                ),
        ),
      ],
    );
  }

  // آئٹم شامل
  void _showAddStockDialog(BuildContext context, BusinessProvider bp) {
    final nameCtrl = TextEditingController();
    final catCtrl = TextEditingController();
    final rateCtrl = TextEditingController();
    final qtyCtrl = TextEditingController(text: '0');
    final alertCtrl = TextEditingController(text: '0');
    String selectedUnit = 'KG';
    final key = GlobalKey<FormState>();

    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => StatefulBuilder(builder: (ctx, setSt) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key, child: SingleChildScrollView(
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const Text('📦 نیا آئٹم شامل کریں', style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
            const SizedBox(height: 16),
            AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
            const SizedBox(height: 12),
            AppFormField(label: '📂 کیٹیگری', controller: catCtrl),
            const SizedBox(height: 12),
            Row(children: [
              Expanded(child: AppFormField(label: '💰 خریداری ریٹ', controller: rateCtrl, required: true,
                  keyboardType: const TextInputType.numberWithOptions(decimal: true))),
              const SizedBox(width: 12),
              Expanded(child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                const Text('یونٹ', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                const SizedBox(height: 6),
                DropdownButtonFormField<String>(
                  value: selectedUnit,
                  items: AppConstants.stockUnits.map((u) => DropdownMenuItem(value: u, child: Text(u))).toList(),
                  onChanged: (v) => setSt(() => selectedUnit = v!),
                  decoration: InputDecoration(border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)), isDense: true, contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10)),
                ),
              ])),
            ]),
            const SizedBox(height: 12),
            Row(children: [
              Expanded(child: AppFormField(label: '📊 ابتدائی مقدار', controller: qtyCtrl,
                  keyboardType: const TextInputType.numberWithOptions(decimal: true))),
              const SizedBox(width: 12),
              Expanded(child: AppFormField(label: '⚠️ الرٹ حد', controller: alertCtrl,
                  keyboardType: const TextInputType.numberWithOptions(decimal: true))),
            ]),
            const SizedBox(height: 20),
            SizedBox(width: double.infinity, child: ElevatedButton(
              onPressed: () async {
                if (!key.currentState!.validate()) return;
                await bp.addStockItem(StockItem(
                  id: 0, name: nameCtrl.text.trim(), category: catCtrl.text.trim(),
                  purchaseRate: double.tryParse(rateCtrl.text) ?? 0,
                  quantity: double.tryParse(qtyCtrl.text) ?? 0,
                  alertLimit: double.tryParse(alertCtrl.text) ?? 0,
                  purchaseHistory: [], createdAt: DateTime.now().toIso8601String(), unit: selectedUnit,
                ));
                if (ctx.mounted) Navigator.pop(ctx);
                if (mounted) showSnackBar(context, '✅ آئٹم شامل!');
              },
              child: const Text('💾 محفوظ'),
            )),
            const SizedBox(height: 16),
          ]),
        )),
      )),
    );
  }

  void _showEditDialog(BuildContext context, BusinessProvider bp, StockItem item) {
    final nameCtrl = TextEditingController(text: item.name);
    final catCtrl = TextEditingController(text: item.category);
    final rateCtrl = TextEditingController(text: item.purchaseRate.toString());
    final alertCtrl = TextEditingController(text: item.alertLimit.toString());
    String selectedUnit = item.unit;
    final key = GlobalKey<FormState>();

    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => StatefulBuilder(builder: (ctx, setSt) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key, child: Column(mainAxisSize: MainAxisSize.min, children: [
          Text('✏️ ${item.name}', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
          const SizedBox(height: 16),
          AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
          const SizedBox(height: 12),
          AppFormField(label: '📂 کیٹیگری', controller: catCtrl),
          const SizedBox(height: 12),
          Row(children: [
            Expanded(child: AppFormField(label: '💰 ریٹ (WAC)', controller: rateCtrl,
                keyboardType: const TextInputType.numberWithOptions(decimal: true))),
            const SizedBox(width: 12),
            Expanded(child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              const Text('یونٹ', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
              const SizedBox(height: 6),
              DropdownButtonFormField<String>(
                value: selectedUnit,
                items: AppConstants.stockUnits.map((u) => DropdownMenuItem(value: u, child: Text(u))).toList(),
                onChanged: (v) => setSt(() => selectedUnit = v!),
                decoration: InputDecoration(border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)), isDense: true, contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10)),
              ),
            ])),
          ]),
          const SizedBox(height: 12),
          AppFormField(label: '⚠️ الرٹ حد', controller: alertCtrl,
              keyboardType: const TextInputType.numberWithOptions(decimal: true)),
          const SizedBox(height: 20),
          SizedBox(width: double.infinity, child: ElevatedButton(
            onPressed: () async {
              if (!key.currentState!.validate()) return;
              await bp.updateStockItem(item.copyWith(
                name: nameCtrl.text.trim(), category: catCtrl.text.trim(),
                purchaseRate: double.tryParse(rateCtrl.text) ?? item.purchaseRate,
                alertLimit: double.tryParse(alertCtrl.text) ?? item.alertLimit, unit: selectedUnit,
              ));
              if (ctx.mounted) Navigator.pop(ctx);
              if (mounted) showSnackBar(context, '✅ محفوظ!');
            },
            child: const Text('💾 محفوظ'),
          )),
          const SizedBox(height: 16),
        ])),
      )),
    );
  }

  void _showAdjDialog(BuildContext context, BusinessProvider bp, StockItem item) {
    String adjType = 'loss';
    final qtyCtrl = TextEditingController();
    final rateCtrl = TextEditingController(text: item.purchaseRate.toStringAsFixed(2));
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();

    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => StatefulBuilder(builder: (ctx, setSt) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key, child: Column(mainAxisSize: MainAxisSize.min, children: [
          Text('⚖️ ${item.name} — ایڈجسٹمنٹ', style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
          const SizedBox(height: 8),
          Container(
            padding: const EdgeInsets.all(10),
            decoration: BoxDecoration(color: AppTheme.info.withOpacity(0.1), borderRadius: BorderRadius.circular(8)),
            child: Text('موجودہ: ${item.quantity} ${item.unit} | ریٹ: Rs. ${item.purchaseRate.toStringAsFixed(2)}',
                style: const TextStyle(fontWeight: FontWeight.bold)),
          ),
          const SizedBox(height: 12),
          SegmentedButton<String>(
            segments: const [
              ButtonSegment(value: 'loss', label: Text('⬇️ نقصان'), icon: Icon(Icons.trending_down)),
              ButtonSegment(value: 'gain', label: Text('⬆️ فائدہ'), icon: Icon(Icons.trending_up)),
            ],
            selected: {adjType},
            onSelectionChanged: (v) => setSt(() => adjType = v.first),
          ),
          const SizedBox(height: 12),
          Row(children: [
            Expanded(child: AppFormField(label: 'مقدار (${item.unit})', controller: qtyCtrl, required: true,
                keyboardType: const TextInputType.numberWithOptions(decimal: true))),
            const SizedBox(width: 12),
            Expanded(child: AppFormField(label: 'ریٹ', controller: rateCtrl,
                keyboardType: const TextInputType.numberWithOptions(decimal: true))),
          ]),
          const SizedBox(height: 12),
          DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
          const SizedBox(height: 12),
          AppFormField(label: '📝 وجہ', controller: noteCtrl),
          const SizedBox(height: 20),
          SizedBox(width: double.infinity, child: ElevatedButton(
            onPressed: () async {
              if (!key.currentState!.validate()) return;
              try {
                await bp.saveStockAdj(
                  itemId: item.id, adjType: adjType,
                  qty: double.parse(qtyCtrl.text),
                  rate: double.tryParse(rateCtrl.text) ?? item.purchaseRate,
                  date: dateCtrl.text, note: noteCtrl.text,
                );
                if (ctx.mounted) Navigator.pop(ctx);
                if (mounted) showSnackBar(context, '✅ ایڈجسٹمنٹ محفوظ!');
              } catch (e) {
                if (mounted) showSnackBar(context, '❌ $e', isError: true);
              }
            },
            style: ElevatedButton.styleFrom(
                backgroundColor: adjType == 'loss' ? AppTheme.danger : AppTheme.success,
                foregroundColor: Colors.white),
            child: const Text('💾 محفوظ'),
          )),
          const SizedBox(height: 16),
        ])),
      )),
    );
  }

  void _showNewAdjDialog(BuildContext context, BusinessProvider bp) {
    if (bp.stockItems.isEmpty) {
      showSnackBar(context, '❌ پہلے آئٹم شامل کریں!', isError: true);
      return;
    }
    StockItem? selectedItem = bp.stockItems.first;
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('آئٹم منتخب کریں'),
        content: DropdownButtonFormField<StockItem>(
          value: selectedItem,
          items: bp.stockItems.map((i) => DropdownMenuItem(value: i, child: Text('${i.name} (${i.quantity} ${i.unit})'))).toList(),
          onChanged: (v) => selectedItem = v,
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('منسوخ')),
          ElevatedButton(
            onPressed: () {
              Navigator.pop(ctx);
              if (selectedItem != null) _showAdjDialog(context, bp, selectedItem!);
            },
            child: const Text('جاری رکھیں'),
          ),
        ],
      ),
    );
  }

  Future<void> _deleteItem(BuildContext context, BusinessProvider bp, StockItem item) async {
    final ok = await showConfirmDialog(context, '🗑️ ${item.name} حذف کریں؟', 'یہ آئٹم اور اس کا تمام تاریخچہ حذف ہو جائے گا!');
    if (ok) {
      await bp.deleteStockItem(item.id);
      if (mounted) showSnackBar(context, '✅ آئٹم حذف!');
    }
  }
}



// ============================================================
// ٹرانزیکشنز اسکرین (فروخت + خریداری)
// ============================================================

class TransactionsScreen extends StatefulWidget {
  const TransactionsScreen({super.key});
  @override
  State<TransactionsScreen> createState() => _TransactionsScreenState();
}

class _TransactionsScreenState extends State<TransactionsScreen>
    with SingleTickerProviderStateMixin {
  late TabController _tabs;
  String _dateFilter = '';

  @override
  void initState() {
    super.initState();
    _tabs = TabController(length: 2, vsync: this);
  }

  @override
  void dispose() {
    _tabs.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final sales = _dateFilter.isEmpty
          ? bp.sales
          : bp.sales.where((s) => s.date == _dateFilter).toList();
      final purchases = _dateFilter.isEmpty
          ? bp.purchases
          : bp.purchases.where((p) => p.date == _dateFilter).toList();

      final totalSales = sales.fold(0.0, (s, x) => s + x.total);
      final totalPurchases = purchases.fold(0.0, (s, x) => s + x.total);

      return Scaffold(
        appBar: AppBar(
          title: const Text('🔄 ٹرانزیکشنز'),
          backgroundColor: const Color(0xFF343a40),
          actions: [
            IconButton(
              icon: const Icon(Icons.calendar_today),
              onPressed: () async {
                final d = await showDatePicker(
                  context: context,
                  initialDate: DateTime.now(),
                  firstDate: DateTime(2000),
                  lastDate: DateTime(2100),
                );
                if (d != null) setState(() => _dateFilter = d.toIso8601String().split('T')[0]);
              },
            ),
            if (_dateFilter.isNotEmpty)
              IconButton(
                icon: const Icon(Icons.clear),
                onPressed: () => setState(() => _dateFilter = ''),
              ),
          ],
          bottom: TabBar(
            controller: _tabs,
            indicatorColor: Colors.white,
            labelColor: Colors.white,
            unselectedLabelColor: Colors.white60,
            tabs: [
              Tab(text: '🛒 فروخت (${sales.length})'),
              Tab(text: '📦 خریداری (${purchases.length})'),
            ],
          ),
        ),
        body: Column(
          children: [
            // فلٹر بار
            if (_dateFilter.isNotEmpty)
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 8),
                color: AppTheme.warning.withOpacity(0.2),
                child: Row(
                  children: [
                    const Icon(Icons.filter_alt, size: 16, color: AppTheme.warning),
                    const SizedBox(width: 6),
                    Text('فلٹر: $_dateFilter', style: const TextStyle(fontWeight: FontWeight.bold)),
                    const Spacer(),
                    TextButton(
                        onPressed: () => setState(() => _dateFilter = ''),
                        child: const Text('صاف کریں', style: TextStyle(color: AppTheme.danger))),
                  ],
                ),
              ),
            // سمری
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 8),
              color: const Color(0xFF343a40),
              child: Row(
                children: [
                  Expanded(child: _chip('فروخت', 'Rs. ${totalSales.toStringAsFixed(0)}', AppTheme.success)),
                  Expanded(child: _chip('خریداری', 'Rs. ${totalPurchases.toStringAsFixed(0)}', AppTheme.warning)),
                ],
              ),
            ),
            Expanded(
              child: TabBarView(
                controller: _tabs,
                children: [
                  _SalesList(sales: sales, bp: bp),
                  _PurchasesList(purchases: purchases, bp: bp),
                ],
              ),
            ),
          ],
        ),
        floatingActionButton: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            FloatingActionButton.extended(
              heroTag: 'purchase',
              onPressed: () => Navigator.push(
                context,
                MaterialPageRoute(builder: (_) => const AddPurchaseScreen()),
              ),
              backgroundColor: AppTheme.warning,
              foregroundColor: Colors.black,
              icon: const Icon(Icons.add_box),
              label: const Text('نئی خریداری'),
            ),
            const SizedBox(height: 10),
            FloatingActionButton.extended(
              heroTag: 'sale',
              onPressed: () => Navigator.push(
                context,
                MaterialPageRoute(builder: (_) => const AddSaleScreen()),
              ),
              backgroundColor: AppTheme.success,
              foregroundColor: Colors.white,
              icon: const Icon(Icons.add_shopping_cart),
              label: const Text('نئی فروخت'),
            ),
          ],
        ),
      );
    });
  }

  Widget _chip(String label, String value, Color color) => Column(children: [
    Text(value, style: TextStyle(color: color, fontWeight: FontWeight.bold)),
    Text(label, style: const TextStyle(color: Colors.white60, fontSize: 11)),
  ]);
}

class _SalesList extends StatelessWidget {
  final List<Sale> sales;
  final BusinessProvider bp;
  const _SalesList({required this.sales, required this.bp});

  @override
  Widget build(BuildContext context) {
    if (sales.isEmpty) {
      return EmptyState(
        message: 'کوئی فروخت نہیں',
        icon: Icons.shopping_cart_outlined,
        buttonLabel: '➕ نئی فروخت',
        onButtonPress: () => Navigator.push(context, MaterialPageRoute(builder: (_) => const AddSaleScreen())),
      );
    }
    return ListView.builder(
      padding: const EdgeInsets.all(12),
      itemCount: sales.length,
      itemBuilder: (ctx, i) {
        final sale = sales[i];
        final customer = bp.getCust(sale.customerId);
        final itemDesc = sale.items.map((si) {
          final item = bp.getStock(si.itemId);
          return '${item?.name ?? '?'} ×${si.qty}${item?.unit ?? ''}';
        }).join('، ');
        final profit = sale.items.fold(0.0, (s, si) {
          final cost = si.costRate > 0 ? si.costRate : bp.getWac(si.itemId);
          return s + si.total - (cost * si.qty);
        });

        return Card(
          margin: const EdgeInsets.only(bottom: 10),
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(customer?.name ?? 'کسٹمر نہیں', style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                        Text(sale.date, style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                        if (itemDesc.isNotEmpty) Text(itemDesc, style: const TextStyle(fontSize: 12)),
                        if (sale.note.isNotEmpty)
                          Text('📝 ${sale.note}', style: const TextStyle(fontSize: 11, color: Colors.grey)),
                      ],
                    )),
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.end,
                      children: [
                        Text('Rs. ${sale.total.toStringAsFixed(2)}',
                            style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.success, fontSize: 16)),
                        if (sale.creditAmount > 0)
                          Text('بقایا: Rs. ${sale.creditAmount.toStringAsFixed(2)}',
                              style: const TextStyle(fontSize: 11, color: AppTheme.danger)),
                        Text('منافع: Rs. ${profit.toStringAsFixed(0)}',
                            style: const TextStyle(fontSize: 11, color: AppTheme.success)),
                      ],
                    ),
                  ],
                ),
                const SizedBox(height: 8),
                Row(
                  children: [
                    if (sale.discount > 0) _tag('ڈسکاؤنٹ: ${sale.discount.toStringAsFixed(0)}', AppTheme.warning),
                    if (sale.fee > 0) _tag('چارج: ${sale.fee.toStringAsFixed(0)}', AppTheme.orange),
                    _tag('نقد: ${sale.cashReceived.toStringAsFixed(0)}', AppTheme.success),
                    const Spacer(),
                    IconButton(
                      icon: const Icon(Icons.delete, color: AppTheme.danger, size: 20),
                      onPressed: () => _deleteSale(context, sale),
                      padding: EdgeInsets.zero,
                    ),
                  ],
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  Widget _tag(String label, Color color) => Container(
    margin: const EdgeInsets.only(right: 6),
    padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 3),
    decoration: BoxDecoration(color: color.withOpacity(0.1), borderRadius: BorderRadius.circular(4)),
    child: Text(label, style: TextStyle(fontSize: 10, color: color, fontWeight: FontWeight.bold)),
  );

  Future<void> _deleteSale(BuildContext context, Sale sale) async {
    final ok = await showConfirmDialog(context, '🗑️ فروخت حذف کریں؟',
        'یہ فروخت حذف ہو گی، اسٹاک واپس آئے گا، اور کسٹمر بیلنس اپڈیٹ ہو گا۔');
    if (ok) {
      await bp.deleteSale(sale.id);
      if (context.mounted) showSnackBar(context, '✅ فروخت حذف!');
    }
  }
}

class _PurchasesList extends StatelessWidget {
  final List<Purchase> purchases;
  final BusinessProvider bp;
  const _PurchasesList({required this.purchases, required this.bp});

  @override
  Widget build(BuildContext context) {
    if (purchases.isEmpty) {
      return EmptyState(
        message: 'کوئی خریداری نہیں',
        icon: Icons.inventory_2_outlined,
        buttonLabel: '➕ نئی خریداری',
        onButtonPress: () => Navigator.push(context, MaterialPageRoute(builder: (_) => const AddPurchaseScreen())),
      );
    }
    return ListView.builder(
      padding: const EdgeInsets.all(12),
      itemCount: purchases.length,
      itemBuilder: (ctx, i) {
        final p = purchases[i];
        final supplier = bp.getSupp(p.supplierId);
        final itemDesc = p.items.map((pi) {
          final item = bp.getStock(pi.itemId);
          return '${item?.name ?? '?'} ×${pi.qty}${item?.unit ?? ''}';
        }).join('، ');

        return Card(
          margin: const EdgeInsets.only(bottom: 10),
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(supplier?.name ?? 'سپلائر نہیں', style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                        Text(p.date, style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                        if (itemDesc.isNotEmpty) Text(itemDesc, style: const TextStyle(fontSize: 12)),
                        if (p.note.isNotEmpty)
                          Text('📝 ${p.note}', style: const TextStyle(fontSize: 11, color: Colors.grey)),
                      ],
                    )),
                    Column(
                      crossAxisAlignment: CrossAxisAlignment.end,
                      children: [
                        Text('Rs. ${p.total.toStringAsFixed(2)}',
                            style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.warning, fontSize: 16)),
                        if (p.creditAmount > 0)
                          Text('بقایا: Rs. ${p.creditAmount.toStringAsFixed(2)}',
                              style: const TextStyle(fontSize: 11, color: AppTheme.danger)),
                      ],
                    ),
                  ],
                ),
                const SizedBox(height: 6),
                Row(
                  children: [
                    if (p.discount > 0) _tag('ڈسکاؤنٹ: ${p.discount.toStringAsFixed(0)}', AppTheme.success),
                    _tag('ادا: ${p.cashPaid.toStringAsFixed(0)}', AppTheme.warning),
                    const Spacer(),
                    IconButton(
                      icon: const Icon(Icons.delete, color: AppTheme.danger, size: 20),
                      onPressed: () => _deletePurchase(context, p),
                      padding: EdgeInsets.zero,
                    ),
                  ],
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  Widget _tag(String label, Color color) => Container(
    margin: const EdgeInsets.only(right: 6),
    padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 3),
    decoration: BoxDecoration(color: color.withOpacity(0.1), borderRadius: BorderRadius.circular(4)),
    child: Text(label, style: TextStyle(fontSize: 10, color: color, fontWeight: FontWeight.bold)),
  );

  Future<void> _deletePurchase(BuildContext context, Purchase p) async {
    final ok = await showConfirmDialog(context, '🗑️ خریداری حذف کریں؟',
        'یہ خریداری حذف ہو گی، اسٹاک واپس کٹے گا، اور سپلائر بیلنس اپڈیٹ ہو گا۔');
    if (ok) {
      await bp.deletePurchase(p.id);
      if (context.mounted) showSnackBar(context, '✅ خریداری حذف!');
    }
  }
}

// ============================================================
// نئی خریداری اسکرین
// ============================================================

class AddPurchaseScreen extends StatefulWidget {
  const AddPurchaseScreen({super.key});
  @override
  State<AddPurchaseScreen> createState() => _AddPurchaseScreenState();
}

class _AddPurchaseScreenState extends State<AddPurchaseScreen> {
  Supplier? _selectedSupplier;
  List<_PurchaseItemEntry> _items = [];
  final _discountCtrl = TextEditingController(text: '0');
  final _feeCtrl = TextEditingController(text: '0');
  final _cashCtrl = TextEditingController(text: '0');
  final _noteCtrl = TextEditingController();
  final _dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
  bool _isProcessing = false;

  @override
  void dispose() {
    _discountCtrl.dispose(); _feeCtrl.dispose();
    _cashCtrl.dispose(); _noteCtrl.dispose(); _dateCtrl.dispose();
    super.dispose();
  }

  double get _subTotal => _items.fold(0.0, (s, i) => s + (i.qty * i.rate));
  double get _discount => double.tryParse(_discountCtrl.text) ?? 0;
  double get _fee => double.tryParse(_feeCtrl.text) ?? 0;
  double get _total => _subTotal - _discount + _fee;
  double get _cashPaid => double.tryParse(_cashCtrl.text) ?? 0;
  double get _creditAmount => _total - _cashPaid;

  void _addItem(StockItem item) {
    final existing = _items.indexWhere((i) => i.itemId == item.id);
    setState(() {
      if (existing != -1) {
        _items[existing].qty += 1;
      } else {
        _items.add(_PurchaseItemEntry(
          itemId: item.id, itemName: item.name, unit: item.unit,
          rate: item.purchaseRate, qty: 1,
        ));
      }
    });
  }

  Future<void> _savePurchase() async {
    if (_selectedSupplier == null) {
      showSnackBar(context, '❌ سپلائر منتخب کریں!', isError: true); return;
    }
    if (_items.isEmpty) {
      showSnackBar(context, '❌ کم از کم ایک آئٹم شامل کریں!', isError: true); return;
    }
    for (final item in _items) {
      if (item.qty <= 0 || item.rate <= 0) {
        showSnackBar(context, '❌ ${item.itemName} کی مقدار/ریٹ درست کریں!', isError: true); return;
      }
    }
    setState(() => _isProcessing = true);
    try {
      final bp = context.read<BusinessProvider>();
      final purchaseItems = _items.map((i) => PurchaseItem(
        itemId: i.itemId, qty: i.qty, rate: i.rate, total: i.qty * i.rate,
      )).toList();
      await bp.addPurchase(
        supplierId: _selectedSupplier!.id,
        items: purchaseItems,
        discount: _discount, fee: _fee,
        cashPaid: _cashPaid > _total ? _total : _cashPaid,
        date: _dateCtrl.text,
        note: _noteCtrl.text.trim().isNotEmpty ? _noteCtrl.text.trim() : '',
      );
      if (mounted) {
        showSnackBar(context, '✅ خریداری مکمل! Rs. ${_total.toStringAsFixed(2)}');
        Navigator.pop(context);
      }
    } catch (e) {
      if (mounted) showSnackBar(context, '❌ $e', isError: true);
    }
    setState(() => _isProcessing = false);
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      return Scaffold(
        appBar: AppBar(
          title: const Text('📦 نئی خریداری'),
          backgroundColor: AppTheme.warning,
          foregroundColor: Colors.black,
        ),
        body: Stack(
          children: [
            SingleChildScrollView(
              padding: const EdgeInsets.all(14),
              child: Column(children: [
                // سپلائر
                Card(child: Padding(
                  padding: const EdgeInsets.all(14),
                  child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                    const Text('🚚 سپلائر منتخب کریں', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16, color: AppTheme.warning)),
                    const SizedBox(height: 12),
                    DropdownButtonFormField<Supplier>(
                      value: _selectedSupplier,
                      items: bp.suppliers.map((s) => DropdownMenuItem(value: s, child: Text(
                        '${s.name}${s.balance > 0 ? ' (بقایا: ${s.balance.toStringAsFixed(0)})' : ''}',
                        style: TextStyle(color: s.balance > 0 ? AppTheme.danger : null),
                      ))).toList(),
                      onChanged: (s) => setState(() => _selectedSupplier = s),
                      decoration: InputDecoration(
                        border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                        hintText: 'سپلائر منتخب کریں',
                      ),
                    ),
                  ]),
                )),
                const SizedBox(height: 12),

                // آئٹمز
                Card(child: Padding(
                  padding: const EdgeInsets.all(14),
                  child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                    const Text('📦 آئٹمز', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16, color: AppTheme.info)),
                    const SizedBox(height: 10),
                    _ItemSearchWidget(stockItems: bp.stockItems, onItemSelected: _addItem),
                    const SizedBox(height: 10),
                    if (_items.isEmpty)
                      const Padding(padding: EdgeInsets.symmetric(vertical: 16),
                        child: Text('کوئی آئٹم نہیں', style: TextStyle(color: Colors.grey))),
                    ..._items.asMap().entries.map((e) => _itemRow(e.key, e.value)),
                  ]),
                )),
                const SizedBox(height: 12),

                // پیمنٹ
                Card(child: Padding(
                  padding: const EdgeInsets.all(14),
                  child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                    const Text('💰 پیمنٹ', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16, color: AppTheme.success)),
                    const SizedBox(height: 10),
                    _sumRow('سب ٹوٹل:', 'Rs. ${_subTotal.toStringAsFixed(2)}'),
                    const SizedBox(height: 8),
                    Row(children: [
                      Expanded(child: _field('ڈسکاؤنٹ', _discountCtrl)),
                      const SizedBox(width: 12),
                      Expanded(child: _field('اضافی چارج', _feeCtrl)),
                    ]),
                    const SizedBox(height: 8),
                    _sumRow('کل:', 'Rs. ${_total.toStringAsFixed(2)}', large: true),
                    const SizedBox(height: 8),
                    _field('نقد ادا (Rs.)', _cashCtrl),
                    const SizedBox(height: 6),
                    Row(mainAxisAlignment: MainAxisAlignment.end, children: [
                      TextButton(onPressed: () { _cashCtrl.text = _total.toStringAsFixed(2); setState(() {}); }, child: const Text('مکمل ادائیگی')),
                      TextButton(onPressed: () { _cashCtrl.text = '0'; setState(() {}); }, child: const Text('بقایا', style: TextStyle(color: AppTheme.danger))),
                    ]),
                    Container(
                      padding: const EdgeInsets.all(10),
                      decoration: BoxDecoration(
                        color: (_creditAmount > 0 ? AppTheme.danger : AppTheme.success).withOpacity(0.08),
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: Row(mainAxisAlignment: MainAxisAlignment.spaceBetween, children: [
                        Text(_creditAmount > 0 ? '⏳ بقایا:' : '✅ مکمل', style: const TextStyle(fontWeight: FontWeight.bold)),
                        Text('Rs. ${_creditAmount.abs().toStringAsFixed(2)}',
                            style: TextStyle(color: _creditAmount > 0 ? AppTheme.danger : AppTheme.success, fontWeight: FontWeight.w900, fontSize: 16)),
                      ]),
                    ),
                    const SizedBox(height: 8),
                    Row(children: [
                      Expanded(child: _dateField()),
                      const SizedBox(width: 12),
                      Expanded(child: _field('نوٹ', _noteCtrl)),
                    ]),
                  ]),
                )),
                const SizedBox(height: 90),
              ]),
            ),
            if (_isProcessing) const LoadingOverlay(),
          ],
        ),
        bottomNavigationBar: Container(
          padding: const EdgeInsets.all(14),
          color: Colors.white,
          child: ElevatedButton(
            onPressed: _isProcessing ? null : _savePurchase,
            style: ElevatedButton.styleFrom(
              backgroundColor: AppTheme.warning, foregroundColor: Colors.black,
              padding: const EdgeInsets.symmetric(vertical: 14),
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
            ),
            child: _isProcessing
                ? const CircularProgressIndicator(color: Colors.black)
                : Text('✅ خریداری مکمل — Rs. ${_total.toStringAsFixed(2)}',
                    style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
          ),
        ),
      );
    });
  }

  Widget _sumRow(String l, String v, {bool large = false}) => Row(
    mainAxisAlignment: MainAxisAlignment.spaceBetween,
    children: [
      Text(l, style: const TextStyle(fontWeight: FontWeight.bold)),
      Text(v, style: TextStyle(fontWeight: large ? FontWeight.w900 : FontWeight.normal,
          fontSize: large ? 16 : 14, color: large ? AppTheme.warning : null)),
    ],
  );

  Widget _field(String label, TextEditingController ctrl) => Column(
    crossAxisAlignment: CrossAxisAlignment.start,
    children: [
      Text(label, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 12)),
      const SizedBox(height: 4),
      TextFormField(
        controller: ctrl,
        keyboardType: const TextInputType.numberWithOptions(decimal: true),
        textAlign: TextAlign.right,
        onChanged: (_) => setState(() {}),
        decoration: InputDecoration(
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
          isDense: true, contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 9),
        ),
      ),
    ],
  );

  Widget _dateField() => Column(
    crossAxisAlignment: CrossAxisAlignment.start,
    children: [
      const Text('📅 تاریخ', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 12)),
      const SizedBox(height: 4),
      TextFormField(
        controller: _dateCtrl,
        readOnly: true,
        onTap: () async {
          final d = await showDatePicker(context: context,
            initialDate: DateTime.now(), firstDate: DateTime(2000), lastDate: DateTime(2100));
          if (d != null) _dateCtrl.text = d.toIso8601String().split('T')[0];
        },
        decoration: InputDecoration(
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
          isDense: true, contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 9),
          suffixIcon: const Icon(Icons.calendar_today, size: 18),
        ),
      ),
    ],
  );

  Widget _itemRow(int idx, _PurchaseItemEntry item) {
    return Card(
      margin: const EdgeInsets.only(bottom: 8),
      color: AppTheme.lightBg,
      child: Padding(
        padding: const EdgeInsets.all(10),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Row(children: [
            Expanded(child: Text(item.itemName, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14))),
            IconButton(
              icon: const Icon(Icons.delete, color: AppTheme.danger, size: 18),
              onPressed: () => setState(() => _items.removeAt(idx)),
              padding: EdgeInsets.zero, constraints: const BoxConstraints(),
            ),
          ]),
          const SizedBox(height: 6),
          Row(children: [
            Expanded(child: _miniInput('مقدار (${item.unit})', item.qty.toString(), (v) {
              setState(() => _items[idx].qty = double.tryParse(v) ?? 0);
            })),
            const SizedBox(width: 10),
            Expanded(child: _miniInput('ریٹ (Rs.)', item.rate.toStringAsFixed(2), (v) {
              setState(() => _items[idx].rate = double.tryParse(v) ?? 0);
            })),
            const SizedBox(width: 10),
            Expanded(child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              const Text('کل', style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold)),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 8),
                decoration: BoxDecoration(color: AppTheme.warning.withOpacity(0.1), borderRadius: BorderRadius.circular(6)),
                child: Text('Rs. ${(item.qty * item.rate).toStringAsFixed(2)}',
                    style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.warning, fontSize: 13)),
              ),
            ])),
          ]),
        ]),
      ),
    );
  }

  Widget _miniInput(String label, String init, ValueChanged<String> onChange) => Column(
    crossAxisAlignment: CrossAxisAlignment.start,
    children: [
      Text(label, style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold)),
      const SizedBox(height: 3),
      TextFormField(
        initialValue: init,
        keyboardType: const TextInputType.numberWithOptions(decimal: true),
        textAlign: TextAlign.right,
        onChanged: onChange,
        decoration: InputDecoration(
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(6)),
          isDense: true, contentPadding: const EdgeInsets.symmetric(horizontal: 8, vertical: 8),
        ),
        style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13),
      ),
    ],
  );
}

class _PurchaseItemEntry {
  final int itemId;
  final String itemName;
  final String unit;
  double rate;
  double qty;

  _PurchaseItemEntry({required this.itemId, required this.itemName, required this.unit, required this.rate, required this.qty});
}



// ============================================================
// نئی فروخت اسکرین — مکمل بزنس لاجک
// ============================================================

class AddSaleScreen extends StatefulWidget {
  const AddSaleScreen({super.key});
  @override
  State<AddSaleScreen> createState() => _AddSaleScreenState();
}

class _AddSaleScreenState extends State<AddSaleScreen> {
  // کسٹمر
  Customer? _selectedCustomer;

  // آئٹمز
  List<_SaleItemEntry> _items = [];

  // پیمنٹ
  final _discountCtrl = TextEditingController(text: '0');
  final _feeCtrl = TextEditingController(text: '0');
  final _cashCtrl = TextEditingController(text: '0');
  final _noteCtrl = TextEditingController();
  final _dateCtrl = TextEditingController(
      text: DateTime.now().toIso8601String().split('T')[0]);

  bool _isProcessing = false;

  @override
  void dispose() {
    _discountCtrl.dispose();
    _feeCtrl.dispose();
    _cashCtrl.dispose();
    _noteCtrl.dispose();
    _dateCtrl.dispose();
    super.dispose();
  }

  // ============================================================
  // کیلکولیشنز
  // ============================================================
  double get _subTotal =>
      _items.fold(0.0, (s, i) => s + (i.qty * i.rate));

  double get _discount =>
      double.tryParse(_discountCtrl.text) ?? 0;

  double get _fee => double.tryParse(_feeCtrl.text) ?? 0;

  double get _total => _subTotal - _discount + _fee;

  double get _cashReceived =>
      double.tryParse(_cashCtrl.text) ?? 0;

  double get _creditAmount => _total - _cashReceived;

  // ============================================================
  // آئٹم شامل
  // ============================================================
  void _addItem(StockItem item) {
    final existing = _items.indexWhere((i) => i.itemId == item.id);
    setState(() {
      if (existing != -1) {
        _items[existing].qty += 1;
      } else {
        _items.add(_SaleItemEntry(
          itemId: item.id,
          itemName: item.name,
          unit: item.unit,
          availableQty: item.quantity,
          rate: item.purchaseRate,
          qty: 1,
          costRate: item.purchaseRate,
        ));
      }
    });
  }

  // ============================================================
  // فروخت محفوظ
  // ============================================================
  Future<void> _saveSale() async {
    if (_selectedCustomer == null) {
      showSnackBar(context, '❌ کسٹمر منتخب کریں!', isError: true);
      return;
    }
    if (_items.isEmpty) {
      showSnackBar(context, '❌ کم از کم ایک آئٹم شامل کریں!', isError: true);
      return;
    }

    // مقدار جانچ
    for (final item in _items) {
      if (item.qty <= 0) {
        showSnackBar(context, '❌ ${item.itemName} کی مقدار درست کریں!', isError: true);
        return;
      }
      if (item.rate <= 0) {
        showSnackBar(context, '❌ ${item.itemName} کا ریٹ درست کریں!', isError: true);
        return;
      }
      if (item.qty > item.availableQty) {
        showSnackBar(context, '❌ ${item.itemName}: صرف ${item.availableQty} ${item.unit} موجود ہے!', isError: true);
        return;
      }
    }

    if (_cashReceived < 0) {
      showSnackBar(context, '❌ نقد رقم منفی نہیں ہو سکتی!', isError: true);
      return;
    }

    setState(() => _isProcessing = true);
    try {
      final bp = context.read<BusinessProvider>();
      final saleItems = _items
          .map((i) => SaleItem(
                itemId: i.itemId,
                qty: i.qty,
                rate: i.rate,
                total: i.qty * i.rate,
                costRate: i.costRate,
              ))
          .toList();

      final saleId = await bp.addSale(
        customerId: _selectedCustomer!.id,
        items: saleItems,
        discount: _discount,
        fee: _fee,
        cashReceived: _cashReceived > _total ? _total : _cashReceived,
        date: _dateCtrl.text,
        note: _noteCtrl.text.trim().isNotEmpty ? _noteCtrl.text.trim() : '',
      );

      if (mounted) {
        _showReceiptDialog(context, bp, saleId);
        Navigator.pop(context);
      }
    } catch (e) {
      if (mounted) showSnackBar(context, '❌ خرابی: $e', isError: true);
    }
    setState(() => _isProcessing = false);
  }

  // ============================================================
  // رسید ڈائیلاگ
  // ============================================================
  void _showReceiptDialog(
      BuildContext context, BusinessProvider bp, int saleId) async {
    final sale = bp.sales.firstWhere((s) => s.id == saleId,
        orElse: () => bp.sales.first);

    if (!mounted) return;
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('✅ فروخت مکمل!',
            style: TextStyle(color: AppTheme.success, fontWeight: FontWeight.bold)),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            _receiptRow('کسٹمر:', _selectedCustomer?.name ?? ''),
            _receiptRow('کل رقم:', 'Rs. ${_total.toStringAsFixed(2)}'),
            if (_discount > 0) _receiptRow('ڈسکاؤنٹ:', 'Rs. ${_discount.toStringAsFixed(2)}'),
            if (_fee > 0) _receiptRow('اضافی چارج:', 'Rs. ${_fee.toStringAsFixed(2)}'),
            _receiptRow('نقد وصول:', 'Rs. ${_cashReceived.toStringAsFixed(2)}'),
            if (_creditAmount > 0)
              _receiptRow('بقایا:', 'Rs. ${_creditAmount.toStringAsFixed(2)}',
                  color: AppTheme.danger),
            if (_creditAmount < 0)
              _receiptRow('واپس کریں:', 'Rs. ${_creditAmount.abs().toStringAsFixed(2)}',
                  color: AppTheme.success),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx),
            child: const Text('✓ بند کریں'),
          ),
          ElevatedButton.icon(
            onPressed: () {
              Navigator.pop(ctx);
              // پرنٹ فیچر
            },
            icon: const Icon(Icons.print, size: 16),
            label: const Text('🖨️ پرنٹ'),
            style: ElevatedButton.styleFrom(backgroundColor: AppTheme.info),
          ),
        ],
      ),
    );
  }

  Widget _receiptRow(String label, String value, {Color? color}) => Padding(
        padding: const EdgeInsets.symmetric(vertical: 3),
        child: Row(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            Text(label, style: const TextStyle(fontWeight: FontWeight.bold)),
            Text(value,
                style: TextStyle(color: color ?? AppTheme.textPrimary, fontWeight: FontWeight.bold)),
          ],
        ),
      );

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      return Scaffold(
        appBar: AppBar(
          title: const Text('🛒 نئی فروخت'),
          backgroundColor: AppTheme.success,
        ),
        body: Stack(
          children: [
            SingleChildScrollView(
              padding: const EdgeInsets.all(14),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  // ============ کسٹمر ============
                  _sectionCard(
                    '👤 کسٹمر منتخب کریں',
                    AppTheme.primary,
                    Column(
                      children: [
                        DropdownButtonFormField<Customer>(
                          value: _selectedCustomer,
                          items: bp.customers
                              .map((c) => DropdownMenuItem(
                                    value: c,
                                    child: Text(
                                      '${c.name}${c.balance > 0 ? ' (بقایا: ${c.balance.toStringAsFixed(0)})' : ''}',
                                      style: TextStyle(
                                          color: c.balance > 0 ? AppTheme.danger : null),
                                    ),
                                  ))
                              .toList(),
                          onChanged: (c) => setState(() => _selectedCustomer = c),
                          decoration: InputDecoration(
                            border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                            hintText: 'کسٹمر منتخب کریں',
                          ),
                        ),
                        if (_selectedCustomer != null) ...[
                          const SizedBox(height: 8),
                          Container(
                            padding: const EdgeInsets.all(8),
                            decoration: BoxDecoration(
                                color: AppTheme.info.withOpacity(0.08),
                                borderRadius: BorderRadius.circular(8)),
                            child: Row(
                              mainAxisAlignment: MainAxisAlignment.spaceBetween,
                              children: [
                                Text('بقایا بیلنس:', style: const TextStyle(fontWeight: FontWeight.bold)),
                                Text(
                                  'Rs. ${_selectedCustomer!.balance.toStringAsFixed(2)}',
                                  style: TextStyle(
                                      color: _selectedCustomer!.balance > 0 ? AppTheme.danger : AppTheme.success,
                                      fontWeight: FontWeight.bold),
                                ),
                              ],
                            ),
                          ),
                        ],
                      ],
                    ),
                  ),
                  const SizedBox(height: 12),

                  // ============ آئٹمز ============
                  _sectionCard(
                    '📦 آئٹمز شامل کریں',
                    AppTheme.info,
                    Column(
                      children: [
                        // آئٹم سرچ
                        _ItemSearchWidget(
                          stockItems: bp.stockItems,
                          onItemSelected: _addItem,
                        ),
                        const SizedBox(height: 10),
                        if (_items.isEmpty)
                          const Padding(
                            padding: EdgeInsets.symmetric(vertical: 16),
                            child: Text('کوئی آئٹم شامل نہیں',
                                style: TextStyle(color: Colors.grey)),
                          )
                        else
                          ..._items.asMap().entries.map((entry) {
                            final idx = entry.key;
                            final item = entry.value;
                            return _itemRow(idx, item, bp);
                          }),
                      ],
                    ),
                  ),
                  const SizedBox(height: 12),

                  // ============ پیمنٹ ============
                  _sectionCard(
                    '💰 پیمنٹ تفصیل',
                    AppTheme.success,
                    Column(
                      children: [
                        // سب ٹوٹل
                        _summaryRow('سب ٹوٹل:', 'Rs. ${_subTotal.toStringAsFixed(2)}'),
                        const SizedBox(height: 10),
                        Row(
                          children: [
                            Expanded(
                              child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  const Text('🏷️ ڈسکاؤنٹ (Rs.)',
                                      style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                                  const SizedBox(height: 4),
                                  TextFormField(
                                    controller: _discountCtrl,
                                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                                    textAlign: TextAlign.right,
                                    onChanged: (_) => setState(() {}),
                                    decoration: InputDecoration(
                                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                      isDense: true,
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10),
                                    ),
                                  ),
                                ],
                              ),
                            ),
                            const SizedBox(width: 12),
                            Expanded(
                              child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  const Text('➕ اضافی چارج (Rs.)',
                                      style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                                  const SizedBox(height: 4),
                                  TextFormField(
                                    controller: _feeCtrl,
                                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                                    textAlign: TextAlign.right,
                                    onChanged: (_) => setState(() {}),
                                    decoration: InputDecoration(
                                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                      isDense: true,
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10),
                                    ),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                        const SizedBox(height: 10),
                        _summaryRow('کل رقم:', 'Rs. ${_total.toStringAsFixed(2)}', bold: true, large: true),
                        const SizedBox(height: 10),
                        Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            const Text('💵 نقد وصول شدہ (Rs.)',
                                style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                            const SizedBox(height: 4),
                            Row(
                              children: [
                                Expanded(
                                  child: TextFormField(
                                    controller: _cashCtrl,
                                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                                    textAlign: TextAlign.right,
                                    onChanged: (_) => setState(() {}),
                                    decoration: InputDecoration(
                                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                      isDense: true,
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10),
                                    ),
                                  ),
                                ),
                                const SizedBox(width: 8),
                                ElevatedButton(
                                  onPressed: () {
                                    _cashCtrl.text = _total.toStringAsFixed(2);
                                    setState(() {});
                                  },
                                  style: ElevatedButton.styleFrom(
                                      backgroundColor: AppTheme.success.withOpacity(0.2),
                                      foregroundColor: AppTheme.success,
                                      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10)),
                                  child: const Text('مکمل', style: TextStyle(fontSize: 12)),
                                ),
                                const SizedBox(width: 6),
                                ElevatedButton(
                                  onPressed: () {
                                    _cashCtrl.text = '0';
                                    setState(() {});
                                  },
                                  style: ElevatedButton.styleFrom(
                                      backgroundColor: AppTheme.danger.withOpacity(0.1),
                                      foregroundColor: AppTheme.danger,
                                      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10)),
                                  child: const Text('ادھار', style: TextStyle(fontSize: 12)),
                                ),
                              ],
                            ),
                          ],
                        ),
                        const SizedBox(height: 10),
                        Container(
                          padding: const EdgeInsets.all(10),
                          decoration: BoxDecoration(
                            color: _creditAmount > 0
                                ? AppTheme.danger.withOpacity(0.08)
                                : _creditAmount < 0
                                    ? AppTheme.success.withOpacity(0.08)
                                    : AppTheme.success.withOpacity(0.05),
                            borderRadius: BorderRadius.circular(8),
                            border: Border.all(
                                color: _creditAmount > 0
                                    ? AppTheme.danger.withOpacity(0.3)
                                    : AppTheme.success.withOpacity(0.3)),
                          ),
                          child: Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              Text(
                                _creditAmount > 0 ? '⏳ بقایا (ادھار):' : _creditAmount < 0 ? '↩️ واپس کریں:' : '✅ مکمل ادائیگی',
                                style: TextStyle(
                                    fontWeight: FontWeight.bold,
                                    color: _creditAmount > 0 ? AppTheme.danger : AppTheme.success),
                              ),
                              Text(
                                'Rs. ${_creditAmount.abs().toStringAsFixed(2)}',
                                style: TextStyle(
                                    color: _creditAmount > 0 ? AppTheme.danger : AppTheme.success,
                                    fontWeight: FontWeight.w900,
                                    fontSize: 16),
                              ),
                            ],
                          ),
                        ),
                        const SizedBox(height: 10),
                        Row(
                          children: [
                            Expanded(
                              child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  const Text('📅 تاریخ', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                                  const SizedBox(height: 4),
                                  TextFormField(
                                    controller: _dateCtrl,
                                    readOnly: true,
                                    onTap: () async {
                                      final d = await showDatePicker(
                                        context: context,
                                        initialDate: DateTime.now(),
                                        firstDate: DateTime(2000),
                                        lastDate: DateTime(2100),
                                      );
                                      if (d != null) {
                                        _dateCtrl.text = d.toIso8601String().split('T')[0];
                                      }
                                    },
                                    decoration: InputDecoration(
                                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                      isDense: true,
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10),
                                      suffixIcon: const Icon(Icons.calendar_today, size: 18),
                                    ),
                                  ),
                                ],
                              ),
                            ),
                            const SizedBox(width: 12),
                            Expanded(
                              child: Column(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  const Text('📝 نوٹ', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                                  const SizedBox(height: 4),
                                  TextFormField(
                                    controller: _noteCtrl,
                                    textAlign: TextAlign.right,
                                    decoration: InputDecoration(
                                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                      isDense: true,
                                      hintText: 'اختیاری',
                                      contentPadding: const EdgeInsets.symmetric(horizontal: 10, vertical: 10),
                                    ),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(height: 90),
                ],
              ),
            ),
            if (_isProcessing) const LoadingOverlay(),
          ],
        ),
        bottomNavigationBar: Container(
          padding: const EdgeInsets.all(14),
          color: Colors.white,
          child: Row(
            children: [
              Expanded(
                child: Container(
                  padding: const EdgeInsets.all(12),
                  decoration: BoxDecoration(
                    color: AppTheme.success.withOpacity(0.1),
                    borderRadius: BorderRadius.circular(10),
                    border: Border.all(color: AppTheme.success.withOpacity(0.3)),
                  ),
                  child: Column(
                    children: [
                      const Text('کل رقم', style: TextStyle(fontSize: 11)),
                      Text('Rs. ${_total.toStringAsFixed(2)}',
                          style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.success, fontSize: 16)),
                    ],
                  ),
                ),
              ),
              const SizedBox(width: 12),
              Expanded(
                flex: 2,
                child: ElevatedButton(
                  onPressed: _isProcessing ? null : _saveSale,
                  style: ElevatedButton.styleFrom(
                    backgroundColor: AppTheme.success,
                    foregroundColor: Colors.white,
                    padding: const EdgeInsets.symmetric(vertical: 14),
                    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                  ),
                  child: _isProcessing
                      ? const CircularProgressIndicator(color: Colors.white)
                      : const Text('✅ فروخت مکمل کریں', style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                ),
              ),
            ],
          ),
        ),
      );
    });
  }

  Widget _sectionCard(String title, Color color, Widget child) => Card(
        margin: EdgeInsets.zero,
        child: Padding(
          padding: const EdgeInsets.all(14),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(title, style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16, color: color)),
              const SizedBox(height: 12),
              child,
            ],
          ),
        ),
      );

  Widget _summaryRow(String label, String value, {bool bold = false, bool large = false}) => Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontWeight: FontWeight.bold)),
          Text(value,
              style: TextStyle(
                  fontWeight: bold ? FontWeight.w900 : FontWeight.normal,
                  fontSize: large ? 18 : 14,
                  color: large ? AppTheme.primary : null)),
        ],
      );

  Widget _itemRow(int idx, _SaleItemEntry item, BusinessProvider bp) {
    return Card(
      margin: const EdgeInsets.only(bottom: 8),
      color: AppTheme.lightBg,
      child: Padding(
        padding: const EdgeInsets.all(10),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                Expanded(
                  child: Text(item.itemName,
                      style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                ),
                Text('دستیاب: ${item.availableQty} ${item.unit}',
                    style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                const SizedBox(width: 8),
                IconButton(
                  icon: const Icon(Icons.delete, color: AppTheme.danger, size: 18),
                  onPressed: () => setState(() => _items.removeAt(idx)),
                  padding: EdgeInsets.zero,
                  constraints: const BoxConstraints(),
                ),
              ],
            ),
            const SizedBox(height: 6),
            Row(
              children: [
                Expanded(child: _miniInput('مقدار (${item.unit})', item.qty.toString(), (v) {
                  setState(() => _items[idx].qty = double.tryParse(v) ?? 0);
                })),
                const SizedBox(width: 10),
                Expanded(child: _miniInput('ریٹ (Rs.)', item.rate.toStringAsFixed(2), (v) {
                  setState(() => _items[idx].rate = double.tryParse(v) ?? 0);
                })),
                const SizedBox(width: 10),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Text('کل', style: TextStyle(fontSize: 11, fontWeight: FontWeight.bold)),
                      Container(
                        padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 8),
                        decoration: BoxDecoration(
                            color: AppTheme.primary.withOpacity(0.1),
                            borderRadius: BorderRadius.circular(6)),
                        child: Text(
                          'Rs. ${(item.qty * item.rate).toStringAsFixed(2)}',
                          style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.primary, fontSize: 13),
                        ),
                      ),
                    ],
                  ),
                ),
              ],
            ),
            if (item.qty > item.availableQty) ...[
              const SizedBox(height: 4),
              Text('⚠️ اسٹاک کم ہے! صرف ${item.availableQty} ${item.unit} دستیاب',
                  style: const TextStyle(color: AppTheme.danger, fontSize: 11)),
            ],
          ],
        ),
      ),
    );
  }

  Widget _miniInput(String label, String initial, ValueChanged<String> onChange) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Text(label, style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold)),
        const SizedBox(height: 3),
        TextFormField(
          initialValue: initial,
          keyboardType: const TextInputType.numberWithOptions(decimal: true),
          textAlign: TextAlign.right,
          onChanged: onChange,
          decoration: InputDecoration(
            border: OutlineInputBorder(borderRadius: BorderRadius.circular(6)),
            isDense: true,
            contentPadding: const EdgeInsets.symmetric(horizontal: 8, vertical: 8),
          ),
          style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13),
        ),
      ],
    );
  }
}

// آئٹم سرچ ویجٹ
class _ItemSearchWidget extends StatefulWidget {
  final List<StockItem> stockItems;
  final ValueChanged<StockItem> onItemSelected;

  const _ItemSearchWidget({required this.stockItems, required this.onItemSelected});

  @override
  State<_ItemSearchWidget> createState() => _ItemSearchWidgetState();
}

class _ItemSearchWidgetState extends State<_ItemSearchWidget> {
  final _ctrl = TextEditingController();
  List<StockItem> _results = [];

  @override
  void dispose() {
    _ctrl.dispose();
    super.dispose();
  }

  void _search(String q) {
    if (q.isEmpty) {
      setState(() => _results = []);
      return;
    }
    setState(() {
      _results = widget.stockItems
          .where((i) =>
              i.name.toLowerCase().contains(q.toLowerCase()) ||
              i.category.toLowerCase().contains(q.toLowerCase()))
          .toList();
    });
  }

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        TextField(
          controller: _ctrl,
          textAlign: TextAlign.right,
          onChanged: _search,
          decoration: InputDecoration(
            hintText: '🔍 آئٹم تلاش کریں...',
            prefixIcon: const Icon(Icons.search, color: AppTheme.info),
            suffixIcon: _ctrl.text.isNotEmpty
                ? IconButton(icon: const Icon(Icons.clear), onPressed: () {
                    _ctrl.clear(); _search('');
                  }) : null,
            border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
            isDense: true,
          ),
        ),
        if (_results.isNotEmpty) ...[
          const SizedBox(height: 8),
          Container(
            constraints: const BoxConstraints(maxHeight: 200),
            decoration: BoxDecoration(
              border: Border.all(color: AppTheme.cardBorder),
              borderRadius: BorderRadius.circular(8),
            ),
            child: ListView.builder(
              shrinkWrap: true,
              itemCount: _results.length,
              itemBuilder: (ctx, i) {
                final item = _results[i];
                return ListTile(
                  dense: true,
                  title: Text(item.name, style: const TextStyle(fontWeight: FontWeight.bold)),
                  subtitle: Text('${item.quantity} ${item.unit} | Rs. ${item.purchaseRate.toStringAsFixed(2)}',
                      style: const TextStyle(fontSize: 11)),
                  trailing: item.isLowStock
                      ? const StatusBadge(label: '⚠️ کم', color: AppTheme.danger)
                      : null,
                  onTap: () {
                    widget.onItemSelected(item);
                    _ctrl.clear();
                    _search('');
                  },
                );
              },
            ),
          ),
        ],
        if (_results.isEmpty && _ctrl.text.isEmpty) ...[
          const SizedBox(height: 8),
          Wrap(
            spacing: 6,
            runSpacing: 6,
            children: widget.stockItems.take(8).map((item) => GestureDetector(
              onTap: () => widget.onItemSelected(item),
              child: Container(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
                decoration: BoxDecoration(
                  color: item.isLowStock ? AppTheme.danger.withOpacity(0.1) : AppTheme.info.withOpacity(0.1),
                  borderRadius: BorderRadius.circular(20),
                  border: Border.all(color: (item.isLowStock ? AppTheme.danger : AppTheme.info).withOpacity(0.3)),
                ),
                child: Text('${item.name} (${item.quantity})',
                    style: TextStyle(fontSize: 12, color: item.isLowStock ? AppTheme.danger : AppTheme.info)),
              ),
            )).toList(),
          ),
        ],
      ],
    );
  }
}

class _SaleItemEntry {
  final int itemId;
  final String itemName;
  final String unit;
  final double availableQty;
  double rate;
  double qty;
  final double costRate;

  _SaleItemEntry({
    required this.itemId,
    required this.itemName,
    required this.unit,
    required this.availableQty,
    required this.rate,
    required this.qty,
    required this.costRate,
  });
}




// ============================================================
// اخراجات اسکرین
// ============================================================

class ExpensesScreen extends StatefulWidget {
  const ExpensesScreen({super.key});
  @override
  State<ExpensesScreen> createState() => _ExpensesScreenState();
}

class _ExpensesScreenState extends State<ExpensesScreen> {
  String _dateFilter = '';
  String _categoryFilter = 'تمام';

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      // فلٹر
      List<Expense> filtered = bp.expenses;
      if (_dateFilter.isNotEmpty) {
        filtered = filtered.where((e) => e.date == _dateFilter).toList();
      }
      if (_categoryFilter != 'تمام') {
        filtered = filtered.where((e) => e.category == _categoryFilter).toList();
      }

      final total = filtered.fold(0.0, (s, e) => s + e.amount);

      // کیٹیگری سمری
      final Map<String, double> catSummary = {};
      for (final e in bp.expenses) {
        catSummary[e.category] = (catSummary[e.category] ?? 0) + e.amount;
      }

      return Scaffold(
        appBar: AppBar(
          title: const Text('💸 اخراجات'),
          backgroundColor: AppTheme.danger,
          actions: [
            IconButton(
              icon: const Icon(Icons.calendar_today),
              onPressed: () async {
                final d = await showDatePicker(
                  context: context,
                  initialDate: DateTime.now(),
                  firstDate: DateTime(2000),
                  lastDate: DateTime(2100),
                );
                if (d != null) {
                  setState(() => _dateFilter = d.toIso8601String().split('T')[0]);
                }
              },
            ),
            if (_dateFilter.isNotEmpty)
              IconButton(
                icon: const Icon(Icons.clear),
                onPressed: () => setState(() => _dateFilter = ''),
              ),
          ],
        ),
        body: Column(
          children: [
            // سمری بار
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
              color: AppTheme.danger,
              child: Row(
                children: [
                  Expanded(child: _chip('کل اخراجات', 'Rs. ${bp.totalExpenses.toStringAsFixed(0)}', Colors.white)),
                  Expanded(child: _chip('آج', 'Rs. ${bp.todayExpTotal.toStringAsFixed(0)}', Colors.orange.shade200)),
                  Expanded(child: _chip('فلٹر', 'Rs. ${total.toStringAsFixed(0)}',
                      _dateFilter.isNotEmpty ? Colors.yellow.shade200 : Colors.white)),
                ],
              ),
            ),

            // فلٹر
            if (_dateFilter.isNotEmpty || _categoryFilter != 'تمام')
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 6),
                color: AppTheme.warning.withOpacity(0.15),
                child: Row(
                  children: [
                    if (_dateFilter.isNotEmpty) ...[
                      const Icon(Icons.filter_alt, size: 14, color: AppTheme.warning),
                      const SizedBox(width: 4),
                      Text('تاریخ: $_dateFilter', style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)),
                    ],
                    if (_categoryFilter != 'تمام') ...[
                      const SizedBox(width: 8),
                      Text('کیٹیگری: $_categoryFilter', style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)),
                    ],
                    const Spacer(),
                    TextButton(
                      onPressed: () => setState(() {
                        _dateFilter = '';
                        _categoryFilter = 'تمام';
                      }),
                      child: const Text('صاف', style: TextStyle(color: AppTheme.danger, fontSize: 12)),
                    ),
                  ],
                ),
              ),

            // کیٹیگری فلٹر چپس
            if (catSummary.isNotEmpty) ...[
              SizedBox(
                height: 44,
                child: ListView(
                  scrollDirection: Axis.horizontal,
                  padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
                  children: [
                    _catChip('تمام', catSummary.values.fold(0.0, (s, v) => s + v)),
                    ...catSummary.entries.map((e) => _catChip(e.key, e.value)),
                  ],
                ),
              ),
            ],

            // لسٹ
            Expanded(
              child: filtered.isEmpty
                  ? EmptyState(
                      message: 'کوئی خرچہ نہیں',
                      icon: Icons.money_off_outlined,
                      buttonLabel: '➕ خرچہ شامل کریں',
                      onButtonPress: () => _showAddDialog(context, bp),
                    )
                  : ListView.builder(
                      padding: const EdgeInsets.all(12),
                      itemCount: filtered.length,
                      itemBuilder: (ctx, i) => _expenseCard(ctx, filtered[i], bp),
                    ),
            ),
          ],
        ),
        floatingActionButton: FloatingActionButton.extended(
          onPressed: () => _showAddDialog(context, bp),
          backgroundColor: AppTheme.danger,
          foregroundColor: Colors.white,
          icon: const Icon(Icons.add),
          label: const Text('خرچہ شامل کریں'),
        ),
      );
    });
  }

  Widget _chip(String label, String value, Color color) => Column(
    children: [
      Text(value, style: TextStyle(color: color, fontWeight: FontWeight.bold, fontSize: 13)),
      Text(label, style: const TextStyle(color: Colors.white60, fontSize: 10)),
    ],
  );

  Widget _catChip(String cat, double total) {
    final isSelected = _categoryFilter == cat;
    return GestureDetector(
      onTap: () => setState(() => _categoryFilter = cat),
      child: Container(
        margin: const EdgeInsets.only(right: 6),
        padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
        decoration: BoxDecoration(
          color: isSelected ? AppTheme.danger : AppTheme.lightBg,
          borderRadius: BorderRadius.circular(20),
          border: Border.all(color: isSelected ? AppTheme.danger : AppTheme.cardBorder),
        ),
        child: Text(
          '$cat (${total.toStringAsFixed(0)})',
          style: TextStyle(
            fontSize: 11,
            color: isSelected ? Colors.white : AppTheme.textPrimary,
            fontWeight: isSelected ? FontWeight.bold : FontWeight.normal,
          ),
        ),
      ),
    );
  }

  Widget _expenseCard(BuildContext context, Expense e, BusinessProvider bp) {
    return Card(
      margin: const EdgeInsets.only(bottom: 8),
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Row(
          children: [
            Container(
              padding: const EdgeInsets.all(10),
              decoration: BoxDecoration(
                color: AppTheme.danger.withOpacity(0.1),
                borderRadius: BorderRadius.circular(10),
              ),
              child: Icon(_getCategoryIcon(e.category), color: AppTheme.danger, size: 22),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(e.category,
                      style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                  Text(e.date, style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                  if (e.note.isNotEmpty)
                    Text(e.note, style: TextStyle(fontSize: 12, color: Colors.grey.shade600)),
                ],
              ),
            ),
            Column(
              crossAxisAlignment: CrossAxisAlignment.end,
              children: [
                Text('Rs. ${e.amount.toStringAsFixed(2)}',
                    style: const TextStyle(
                        color: AppTheme.danger, fontWeight: FontWeight.w900, fontSize: 16)),
                Row(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    IconButton(
                      icon: const Icon(Icons.edit, size: 16, color: AppTheme.orange),
                      onPressed: () => _showEditDialog(context, bp, e),
                      padding: EdgeInsets.zero,
                      constraints: const BoxConstraints(),
                    ),
                    const SizedBox(width: 4),
                    IconButton(
                      icon: const Icon(Icons.delete, size: 16, color: AppTheme.danger),
                      onPressed: () => _deleteExpense(context, bp, e),
                      padding: EdgeInsets.zero,
                      constraints: const BoxConstraints(),
                    ),
                  ],
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }

  IconData _getCategoryIcon(String cat) {
    switch (cat) {
      case 'کرایہ': return Icons.home;
      case 'بجلی': return Icons.bolt;
      case 'پانی': return Icons.water_drop;
      case 'گیس': return Icons.local_fire_department;
      case 'تنخواہ': return Icons.person;
      case 'ٹرانسپورٹ': return Icons.local_shipping;
      case 'مرمت': return Icons.build;
      case 'ٹیلیفون': return Icons.phone;
      case 'انٹرنیٹ': return Icons.wifi;
      default: return Icons.receipt_long;
    }
  }

  void _showAddDialog(BuildContext context, BusinessProvider bp) {
    String selectedCategory = expenseCategories.first;
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(
        text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();

    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setSt) => Padding(
          padding: EdgeInsets.only(
              bottom: MediaQuery.of(ctx).viewInsets.bottom,
              left: 20, right: 20, top: 20),
          child: Form(
            key: key,
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                const Text('💸 نیا خرچہ',
                    style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
                const SizedBox(height: 16),

                // کیٹیگری
                const Text('📂 کیٹیگری *',
                    style: TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                const SizedBox(height: 6),
                DropdownButtonFormField<String>(
                  value: selectedCategory,
                  items: expenseCategories
                      .map((c) => DropdownMenuItem(value: c, child: Text(c)))
                      .toList(),
                  onChanged: (v) => setSt(() => selectedCategory = v!),
                  decoration: InputDecoration(
                    border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                  ),
                ),
                const SizedBox(height: 12),
                AppFormField(
                  label: '💰 رقم',
                  controller: amountCtrl,
                  required: true,
                  keyboardType:
                      const TextInputType.numberWithOptions(decimal: true),
                  validator: (v) => (v == null ||
                          v.isEmpty ||
                          (double.tryParse(v) ?? 0) <= 0)
                      ? 'درست رقم درج کریں'
                      : null,
                ),
                const SizedBox(height: 12),
                DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
                const SizedBox(height: 12),
                AppFormField(
                    label: '📝 نوٹ', controller: noteCtrl,
                    hint: 'اختیاری تفصیل'),
                const SizedBox(height: 20),
                SizedBox(
                  width: double.infinity,
                  child: ElevatedButton(
                    onPressed: () async {
                      if (!key.currentState!.validate()) return;
                      await bp.addExpense(
                        category: selectedCategory,
                        amount: double.parse(amountCtrl.text),
                        date: dateCtrl.text,
                        note: noteCtrl.text.trim().isNotEmpty
                            ? noteCtrl.text.trim()
                            : '',
                      );
                      if (ctx.mounted) Navigator.pop(ctx);
                      if (mounted) showSnackBar(context, '✅ خرچہ محفوظ!');
                    },
                    style: ElevatedButton.styleFrom(
                        backgroundColor: AppTheme.danger,
                        foregroundColor: Colors.white),
                    child: const Text('💾 محفوظ کریں'),
                  ),
                ),
                const SizedBox(height: 16),
              ],
            ),
          ),
        ),
      ),
    );
  }

  void _showEditDialog(BuildContext context, BusinessProvider bp, Expense e) {
    String selectedCategory = e.category;
    final amountCtrl = TextEditingController(text: e.amount.toString());
    final noteCtrl = TextEditingController(text: e.note ?? '');
    final dateCtrl = TextEditingController(text: e.date);
    final key = GlobalKey<FormState>();

    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(
          borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setSt) => Padding(
          padding: EdgeInsets.only(
              bottom: MediaQuery.of(ctx).viewInsets.bottom,
              left: 20, right: 20, top: 20),
          child: Form(
            key: key,
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                const Text('✏️ خرچہ ترمیم',
                    style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
                const SizedBox(height: 16),
                const Text('📂 کیٹیگری',
                    style: TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                const SizedBox(height: 6),
                DropdownButtonFormField<String>(
                  value: selectedCategory,
                  items: expenseCategories
                      .map((c) => DropdownMenuItem(value: c, child: Text(c)))
                      .toList(),
                  onChanged: (v) => setSt(() => selectedCategory = v!),
                  decoration: InputDecoration(
                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8))),
                ),
                const SizedBox(height: 12),
                AppFormField(
                    label: '💰 رقم', controller: amountCtrl, required: true,
                    keyboardType: const TextInputType.numberWithOptions(decimal: true)),
                const SizedBox(height: 12),
                DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
                const SizedBox(height: 12),
                AppFormField(label: '📝 نوٹ', controller: noteCtrl),
                const SizedBox(height: 20),
                SizedBox(
                  width: double.infinity,
                  child: ElevatedButton(
                    onPressed: () async {
                      if (!key.currentState!.validate()) return;
                      final cashDiff =
                          (double.tryParse(amountCtrl.text) ?? 0) - e.amount;
                      // پہلے پرانا حذف، پھر نیا شامل کریں
                      await bp.deleteExpense(e.id);
                      await bp.addExpense(
                        category: selectedCategory,
                        amount: double.parse(amountCtrl.text),
                        date: dateCtrl.text,
                        note: noteCtrl.text.trim().isNotEmpty ? noteCtrl.text.trim() : '',
                      );
                      if (ctx.mounted) Navigator.pop(ctx);
                      if (mounted) showSnackBar(context, '✅ ترمیم محفوظ!');
                    },
                    style: ElevatedButton.styleFrom(
                        backgroundColor: AppTheme.orange, foregroundColor: Colors.white),
                    child: const Text('💾 محفوظ'),
                  ),
                ),
                const SizedBox(height: 16),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Future<void> _deleteExpense(
      BuildContext context, BusinessProvider bp, Expense e) async {
    final ok = await showConfirmDialog(
        context, '🗑️ خرچہ حذف کریں؟',
        'Rs. ${e.amount.toStringAsFixed(2)} — ${e.category} — ${e.date}');
    if (ok) {
      await bp.deleteExpense(e.id);
      if (mounted) showSnackBar(context, '✅ خرچہ حذف!');
    }
  }

  static const List<String> expenseCategories = [
    'کرایہ', 'بجلی', 'پانی', 'گیس', 'تنخواہ', 'ٹرانسپورٹ',
    'مرمت', 'اشتہار', 'ٹیکس', 'بیمہ', 'دفتری اخراجات', 'صفائی',
    'سیکیورٹی', 'ٹیلیفون', 'انٹرنیٹ', 'پیکنگ', 'لوڈنگ/انلوڈنگ', 'دیگر',
  ];
}




// ============================================================
// رپورٹس اسکرین — تاریخ کے مطابق مکمل رپورٹ
// ============================================================

class ReportsScreen extends StatefulWidget {
  const ReportsScreen({super.key});
  @override
  State<ReportsScreen> createState() => _ReportsScreenState();
}

class _ReportsScreenState extends State<ReportsScreen>
    with SingleTickerProviderStateMixin {
  late TabController _tabs;

  // تاریخ فلٹر
  String _fromDate = '';
  String _toDate = '';

  @override
  void initState() {
    super.initState();
    _tabs = TabController(length: 4, vsync: this);
    // آج کی تاریخ ڈیفالٹ
    final today = DateTime.now().toIso8601String().split('T')[0];
    _fromDate = today;
    _toDate = today;
  }

  @override
  void dispose() {
    _tabs.dispose();
    super.dispose();
  }

  Future<void> _pickDateRange() async {
    final range = await showDateRangePicker(
      context: context,
      firstDate: DateTime(2020),
      lastDate: DateTime(2100),
      initialDateRange: DateTimeRange(
        start: DateTime.tryParse(_fromDate) ?? DateTime.now(),
        end: DateTime.tryParse(_toDate) ?? DateTime.now(),
      ),
      builder: (ctx, child) => Theme(
        data: Theme.of(ctx).copyWith(
          colorScheme: const ColorScheme.light(primary: AppTheme.primary),
        ),
        child: child!,
      ),
    );
    if (range != null) {
      setState(() {
        _fromDate = range.start.toIso8601String().split('T')[0];
        _toDate = range.end.toIso8601String().split('T')[0];
      });
    }
  }

  void _setPreset(String preset) {
    final now = DateTime.now();
    setState(() {
      switch (preset) {
        case 'آج':
          _fromDate = _toDate = now.toIso8601String().split('T')[0];
          break;
        case 'کل':
          final y = now.subtract(const Duration(days: 1));
          _fromDate = _toDate = y.toIso8601String().split('T')[0];
          break;
        case 'اس ہفتے':
          final start = now.subtract(Duration(days: now.weekday - 1));
          _fromDate = start.toIso8601String().split('T')[0];
          _toDate = now.toIso8601String().split('T')[0];
          break;
        case 'اس مہینے':
          _fromDate = DateTime(now.year, now.month, 1).toIso8601String().split('T')[0];
          _toDate = now.toIso8601String().split('T')[0];
          break;
        case 'اس سال':
          _fromDate = DateTime(now.year, 1, 1).toIso8601String().split('T')[0];
          _toDate = now.toIso8601String().split('T')[0];
          break;
        case 'تمام':
          _fromDate = '2000-01-01';
          _toDate = '2099-12-31';
          break;
      }
    });
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final sales = bp.getSalesByDateRange(_fromDate, _toDate);
      final purchases = bp.getPurchasesByDateRange(_fromDate, _toDate);
      final expenses = bp.getExpensesByDateRange(_fromDate, _toDate);

      final totalSales = sales.fold(0.0, (s, x) => s + x.total);
      final totalPurchases = purchases.fold(0.0, (s, x) => s + x.total);
      final totalExpenses = expenses.fold(0.0, (s, x) => s + x.amount);
      final totalCashIn = sales.fold(0.0, (s, x) => s + x.cashReceived);
      final totalCashOut = purchases.fold(0.0, (s, x) => s + x.cashPaid) + totalExpenses;

      double grossProfit = 0;
      for (final sale in sales) {
        for (final item in sale.items) {
          final cost = item.costRate > 0
              ? item.costRate
              : bp.getWac(item.itemId);
          grossProfit += item.total - (cost * item.qty);
        }
      }
      final netProfit = grossProfit - totalExpenses;

      return Scaffold(
        appBar: AppBar(
          title: const Text('📊 رپورٹس'),
          backgroundColor: AppTheme.purple,
          actions: [
            IconButton(
              icon: const Icon(Icons.date_range),
              onPressed: _pickDateRange,
              tooltip: 'تاریخ منتخب کریں',
            ),
          ],
          bottom: TabBar(
            controller: _tabs,
            indicatorColor: Colors.white,
            labelColor: Colors.white,
            unselectedLabelColor: Colors.white60,
            isScrollable: true,
            tabs: const [
              Tab(text: '📋 خلاصہ'),
              Tab(text: '🛒 فروخت'),
              Tab(text: '📦 خریداری'),
              Tab(text: '💸 اخراجات'),
            ],
          ),
        ),
        body: Column(
          children: [
            // تاریخ فلٹر بار
            Container(
              padding: const EdgeInsets.all(10),
              color: AppTheme.purple.withOpacity(0.9),
              child: Column(
                children: [
                  // پری سیٹ بٹن
                  SizedBox(
                    height: 34,
                    child: ListView(
                      scrollDirection: Axis.horizontal,
                      children: ['آج', 'کل', 'اس ہفتے', 'اس مہینے', 'اس سال', 'تمام']
                          .map((p) => GestureDetector(
                                onTap: () => _setPreset(p),
                                child: Container(
                                  margin: const EdgeInsets.only(right: 6),
                                  padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
                                  decoration: BoxDecoration(
                                    color: Colors.white.withOpacity(0.15),
                                    borderRadius: BorderRadius.circular(20),
                                    border: Border.all(color: Colors.white30),
                                  ),
                                  child: Text(p,
                                      style: const TextStyle(color: Colors.white, fontSize: 12)),
                                ),
                              ))
                          .toList(),
                    ),
                  ),
                  const SizedBox(height: 6),
                  GestureDetector(
                    onTap: _pickDateRange,
                    child: Container(
                      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
                      decoration: BoxDecoration(
                        color: Colors.white,
                        borderRadius: BorderRadius.circular(8),
                      ),
                      child: Row(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          const Icon(Icons.date_range, size: 16, color: AppTheme.purple),
                          const SizedBox(width: 6),
                          Text(
                            '$_fromDate  ←  $_toDate',
                            style: const TextStyle(
                                color: AppTheme.purple, fontWeight: FontWeight.bold, fontSize: 13),
                          ),
                        ],
                      ),
                    ),
                  ),
                ],
              ),
            ),

            Expanded(
              child: TabBarView(
                controller: _tabs,
                children: [
                  // ====== خلاصہ ======
                  _buildSummaryTab(
                    sales: sales, purchases: purchases, expenses: expenses,
                    totalSales: totalSales, totalPurchases: totalPurchases,
                    totalExpenses: totalExpenses, totalCashIn: totalCashIn,
                    totalCashOut: totalCashOut, grossProfit: grossProfit,
                    netProfit: netProfit, bp: bp,
                  ),

                  // ====== فروخت ======
                  _buildSalesTab(sales, bp),

                  // ====== خریداری ======
                  _buildPurchasesTab(purchases, bp),

                  // ====== اخراجات ======
                  _buildExpensesTab(expenses),
                ],
              ),
            ),
          ],
        ),
      );
    });
  }

  // ====== خلاصہ ٹیب ======
  Widget _buildSummaryTab({
    required List<Sale> sales,
    required List<Purchase> purchases,
    required List<Expense> expenses,
    required double totalSales,
    required double totalPurchases,
    required double totalExpenses,
    required double totalCashIn,
    required double totalCashOut,
    required double grossProfit,
    required double netProfit,
    required BusinessProvider bp,
  }) {
    return SingleChildScrollView(
      padding: const EdgeInsets.all(14),
      child: Column(
        children: [
          // مین سمری گرڈ
          GridView.count(
            shrinkWrap: true,
            physics: const NeverScrollableScrollPhysics(),
            crossAxisCount: 2,
            crossAxisSpacing: 10,
            mainAxisSpacing: 10,
            childAspectRatio: 1.4,
            children: [
              SummaryCard(
                title: '🛒 کل فروخت',
                amount: 'Rs. ${totalSales.toStringAsFixed(0)}',
                icon: Icons.shopping_cart,
                color: AppTheme.success,
                bgColor: AppTheme.success.withOpacity(0.08),
                subtitle: '${sales.length} ٹرانزیکشن',
              ),
              SummaryCard(
                title: '📦 کل خریداری',
                amount: 'Rs. ${totalPurchases.toStringAsFixed(0)}',
                icon: Icons.inventory_2,
                color: AppTheme.warning,
                bgColor: AppTheme.warning.withOpacity(0.08),
                subtitle: '${purchases.length} ٹرانزیکشن',
              ),
              SummaryCard(
                title: '💸 کل اخراجات',
                amount: 'Rs. ${totalExpenses.toStringAsFixed(0)}',
                icon: Icons.money_off,
                color: AppTheme.danger,
                bgColor: AppTheme.danger.withOpacity(0.08),
                subtitle: '${expenses.length} اندراج',
              ),
              SummaryCard(
                title: '📈 خالص منافع',
                amount: 'Rs. ${netProfit.toStringAsFixed(0)}',
                icon: Icons.trending_up,
                color: netProfit >= 0 ? AppTheme.success : AppTheme.danger,
                bgColor: (netProfit >= 0 ? AppTheme.success : AppTheme.danger).withOpacity(0.08),
              ),
            ],
          ),
          const SizedBox(height: 14),

          // کیش فلو
          _reportSection(
            '💰 کیش فلو',
            AppTheme.primary,
            [
              _reportRow('کل نقد وصول', 'Rs. ${totalCashIn.toStringAsFixed(2)}', AppTheme.success),
              _reportRow('کل نقد ادائیگی', 'Rs. ${totalCashOut.toStringAsFixed(2)}', AppTheme.danger),
              _reportRow('خالص کیش فلو',
                  'Rs. ${(totalCashIn - totalCashOut).toStringAsFixed(2)}',
                  (totalCashIn - totalCashOut) >= 0 ? AppTheme.success : AppTheme.danger,
                  bold: true),
            ],
          ),
          const SizedBox(height: 12),

          // منافع تجزیہ
          _reportSection(
            '📈 منافع تجزیہ',
            AppTheme.success,
            [
              _reportRow('مجموعی منافع', 'Rs. ${grossProfit.toStringAsFixed(2)}', AppTheme.success),
              _reportRow('اخراجات', '- Rs. ${totalExpenses.toStringAsFixed(2)}', AppTheme.danger),
              const Divider(),
              _reportRow('خالص منافع', 'Rs. ${netProfit.toStringAsFixed(2)}',
                  netProfit >= 0 ? AppTheme.success : AppTheme.danger, bold: true),
            ],
          ),
          const SizedBox(height: 12),

          // کیٹیگری اخراجات
          if (expenses.isNotEmpty) ...[
            _buildExpenseCategorySummary(expenses),
            const SizedBox(height: 12),
          ],

          // آئٹم فروخت خلاصہ
          _buildItemSalesSummary(sales, bp),
        ],
      ),
    );
  }

  Widget _buildExpenseCategorySummary(List<Expense> expenses) {
    final Map<String, double> cats = {};
    for (final e in expenses) {
      cats[e.category] = (cats[e.category] ?? 0) + e.amount;
    }
    final sorted = cats.entries.toList()
      ..sort((a, b) => b.value.compareTo(a.value));

    return _reportSection(
      '📂 اخراجات بلحاظ کیٹیگری',
      AppTheme.danger,
      sorted.map((e) => _reportRow(e.key, 'Rs. ${e.value.toStringAsFixed(0)}', AppTheme.danger)).toList(),
    );
  }

  Widget _buildItemSalesSummary(List<Sale> sales, BusinessProvider bp) {
    final Map<int, double> itemQty = {};
    final Map<int, double> itemRevenue = {};
    final Map<int, double> itemProfit = {};

    for (final sale in sales) {
      for (final si in sale.items) {
        itemQty[si.itemId] = (itemQty[si.itemId] ?? 0) + si.qty;
        itemRevenue[si.itemId] = (itemRevenue[si.itemId] ?? 0) + si.total;
        final cost = si.costRate > 0 ? si.costRate : bp.getWac(si.itemId);
        itemProfit[si.itemId] = (itemProfit[si.itemId] ?? 0) + (si.total - cost * si.qty);
      }
    }

    if (itemRevenue.isEmpty) return const SizedBox();

    final sorted = itemRevenue.entries.toList()
      ..sort((a, b) => b.value.compareTo(a.value));

    return Card(
      child: Column(
        children: [
          Container(
            padding: const EdgeInsets.all(12),
            decoration: const BoxDecoration(
              color: AppTheme.info,
              borderRadius: BorderRadius.only(topLeft: Radius.circular(12), topRight: Radius.circular(12)),
            ),
            child: const Row(
              children: [
                Icon(Icons.bar_chart, color: Colors.white),
                SizedBox(width: 8),
                Text('📦 آئٹم وار فروخت', style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 15)),
              ],
            ),
          ),
          ...sorted.take(10).map((entry) {
            final item = bp.getStock(entry.key);
            return Container(
              padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
              decoration: BoxDecoration(
                  border: Border(bottom: BorderSide(color: AppTheme.cardBorder.withOpacity(0.4)))),
              child: Row(
                children: [
                  Expanded(
                    child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                      Text(item?.name ?? '?', style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                      Text('${itemQty[entry.key]?.toStringAsFixed(1)} ${item?.unit ?? ''}',
                          style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                    ]),
                  ),
                  Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                    Text('Rs. ${entry.value.toStringAsFixed(0)}',
                        style: const TextStyle(fontWeight: FontWeight.bold, color: AppTheme.info, fontSize: 13)),
                    Text('منافع: Rs. ${(itemProfit[entry.key] ?? 0).toStringAsFixed(0)}',
                        style: const TextStyle(fontSize: 10, color: AppTheme.success)),
                  ]),
                ],
              ),
            );
          }),
        ],
      ),
    );
  }

  // ====== فروخت ٹیب ======
  Widget _buildSalesTab(List<Sale> sales, BusinessProvider bp) {
    if (sales.isEmpty) {
      return const EmptyState(message: 'اس مدت میں کوئی فروخت نہیں', icon: Icons.shopping_cart_outlined);
    }
    return ListView.builder(
      padding: const EdgeInsets.all(12),
      itemCount: sales.length,
      itemBuilder: (ctx, i) {
        final sale = sales[i];
        final customer = bp.getCust(sale.customerId);
        final itemDesc = sale.items.map((si) {
          final item = bp.getStock(si.itemId);
          return '${item?.name ?? '?'} ×${si.qty}${item?.unit ?? ''}@${si.rate.toStringAsFixed(0)}';
        }).join('، ');
        double saleProfit = 0;
        for (final si in sale.items) {
          final cost = si.costRate > 0 ? si.costRate : bp.getWac(si.itemId);
          saleProfit += si.total - (cost * si.qty);
        }
        return Card(
          margin: const EdgeInsets.only(bottom: 8),
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              Row(children: [
                Expanded(child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                  Text(customer?.name ?? '-', style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                  Text(sale.date, style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                  Text(itemDesc, style: const TextStyle(fontSize: 11)),
                ])),
                Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                  Text('Rs. ${sale.total.toStringAsFixed(2)}',
                      style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.success, fontSize: 15)),
                  if (sale.creditAmount > 0)
                    Text('بقایا: ${sale.creditAmount.toStringAsFixed(0)}',
                        style: const TextStyle(fontSize: 10, color: AppTheme.danger)),
                  Text('منافع: Rs. ${saleProfit.toStringAsFixed(0)}',
                      style: const TextStyle(fontSize: 10, color: AppTheme.success)),
                ]),
              ]),
            ]),
          ),
        );
      },
    );
  }

  // ====== خریداری ٹیب ======
  Widget _buildPurchasesTab(List<Purchase> purchases, BusinessProvider bp) {
    if (purchases.isEmpty) {
      return const EmptyState(message: 'اس مدت میں کوئی خریداری نہیں', icon: Icons.inventory_2_outlined);
    }
    return ListView.builder(
      padding: const EdgeInsets.all(12),
      itemCount: purchases.length,
      itemBuilder: (ctx, i) {
        final p = purchases[i];
        final supplier = bp.getSupp(p.supplierId);
        final itemDesc = p.items.map((pi) {
          final item = bp.getStock(pi.itemId);
          return '${item?.name ?? '?'} ×${pi.qty}${item?.unit ?? ''}@${pi.rate.toStringAsFixed(0)}';
        }).join('، ');
        return Card(
          margin: const EdgeInsets.only(bottom: 8),
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Row(children: [
              Expanded(child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                Text(supplier?.name ?? '-', style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
                Text(p.date, style: TextStyle(fontSize: 11, color: Colors.grey.shade500)),
                Text(itemDesc, style: const TextStyle(fontSize: 11)),
              ])),
              Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                Text('Rs. ${p.total.toStringAsFixed(2)}',
                    style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.warning, fontSize: 15)),
                if (p.creditAmount > 0)
                  Text('بقایا: ${p.creditAmount.toStringAsFixed(0)}',
                      style: const TextStyle(fontSize: 10, color: AppTheme.danger)),
                Text('ادا: ${p.cashPaid.toStringAsFixed(0)}',
                    style: const TextStyle(fontSize: 10, color: AppTheme.success)),
              ]),
            ]),
          ),
        );
      },
    );
  }

  // ====== اخراجات ٹیب ======
  Widget _buildExpensesTab(List<Expense> expenses) {
    if (expenses.isEmpty) {
      return const EmptyState(message: 'اس مدت میں کوئی خرچہ نہیں', icon: Icons.money_off_outlined);
    }
    final total = expenses.fold(0.0, (s, e) => s + e.amount);
    return ListView(
      padding: const EdgeInsets.all(12),
      children: [
        Container(
          margin: const EdgeInsets.only(bottom: 12),
          padding: const EdgeInsets.all(12),
          decoration: BoxDecoration(
            color: AppTheme.danger.withOpacity(0.1),
            borderRadius: BorderRadius.circular(10),
            border: Border.all(color: AppTheme.danger.withOpacity(0.3)),
          ),
          child: Row(mainAxisAlignment: MainAxisAlignment.spaceBetween, children: [
            const Text('کل اخراجات:', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
            Text('Rs. ${total.toStringAsFixed(2)}',
                style: const TextStyle(fontWeight: FontWeight.w900, color: AppTheme.danger, fontSize: 18)),
          ]),
        ),
        ...expenses.map((e) => Card(
          margin: const EdgeInsets.only(bottom: 8),
          child: ListTile(
            leading: CircleAvatar(
              backgroundColor: AppTheme.danger.withOpacity(0.1),
              child: const Icon(Icons.receipt_long, color: AppTheme.danger, size: 18),
            ),
            title: Text(e.category, style: const TextStyle(fontWeight: FontWeight.bold)),
            subtitle: Text('${e.date}${e.note.isNotEmpty ? " | ${e.note}" : ""}',
                style: const TextStyle(fontSize: 11)),
            trailing: Text('Rs. ${e.amount.toStringAsFixed(2)}',
                style: const TextStyle(fontWeight: FontWeight.bold, color: AppTheme.danger)),
          ),
        )),
      ],
    );
  }

  // ====== ہیلپرز ======
  Widget _reportSection(String title, Color color, List<Widget> rows) {
    return Card(
      child: Column(
        children: [
          Container(
            width: double.infinity,
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
            decoration: BoxDecoration(
              color: color,
              borderRadius: const BorderRadius.only(topLeft: Radius.circular(12), topRight: Radius.circular(12)),
            ),
            child: Text(title, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 15)),
          ),
          Padding(
            padding: const EdgeInsets.all(12),
            child: Column(children: rows),
          ),
        ],
      ),
    );
  }

  Widget _reportRow(String label, String value, Color color, {bool bold = false}) {
    if (label == '__divider__') return const Divider();
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 13)),
          Text(value,
              style: TextStyle(
                  color: color,
                  fontWeight: bold ? FontWeight.w900 : FontWeight.bold,
                  fontSize: bold ? 15 : 13)),
        ],
      ),
    );
  }
}


// ============================================================
// منافع تجزیہ اسکرین
// ============================================================

class ProfitScreen extends StatefulWidget {
  const ProfitScreen({super.key});
  @override
  State<ProfitScreen> createState() => _ProfitScreenState();
}

class _ProfitScreenState extends State<ProfitScreen>
    with SingleTickerProviderStateMixin {
  late TabController _tabs;
  int _selectedYear = DateTime.now().year;

  @override
  void initState() {
    super.initState();
    _tabs = TabController(length: 3, vsync: this);
  }

  @override
  void dispose() {
    _tabs.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      return Scaffold(
        appBar: AppBar(
          title: const Text('📈 منافع تجزیہ'),
          backgroundColor: AppTheme.success,
          bottom: TabBar(
            controller: _tabs,
            indicatorColor: Colors.white,
            labelColor: Colors.white,
            unselectedLabelColor: Colors.white60,
            tabs: const [
              Tab(text: 'مجموعی'),
              Tab(text: 'ماہانہ'),
              Tab(text: 'آئٹم وار'),
            ],
          ),
        ),
        body: TabBarView(
          controller: _tabs,
          children: [
            _buildOverallTab(bp),
            _buildMonthlyTab(bp),
            _buildItemWiseTab(bp),
          ],
        ),
      );
    });
  }

  // ====== مجموعی ٹیب ======
  Widget _buildOverallTab(BusinessProvider bp) {
    return SingleChildScrollView(
      padding: const EdgeInsets.all(14),
      child: Column(
        children: [
          // خالص منافع بڑا باکس
          GradientSummaryBox(
            title: '📈 کل خالص منافع',
            amount: 'Rs. ${bp.netProfit.toStringAsFixed(2)}',
            icon: Icons.trending_up,
            colors: bp.netProfit >= 0
                ? [AppTheme.success, const Color(0xFF20c997)]
                : [AppTheme.danger, const Color(0xFFe74c3c)],
            subtitle: bp.netProfit >= 0 ? 'فائدے میں ہیں ✅' : 'نقصان میں ہیں ❌',
          ),
          const SizedBox(height: 14),

          // منافع تجزیہ بریک ڈاؤن
          _profitCard(
            '📊 منافع حساب',
            [
              _row('مجموعی منافع (فروخت)', bp.totalGrossProfit, AppTheme.success),
              _divider(),
              _row('اخراجات', -bp.totalExpenses, AppTheme.danger),
              _row('وزن نقصان', -bp.totalAdjLoss, AppTheme.danger),
              _row('وزن فائدہ', bp.totalAdjGain, AppTheme.success),
              _row('کسٹمر ڈسکاؤنٹ', -bp.totalCustDisc, AppTheme.warning),
              _row('سپلائر ڈسکاؤنٹ', bp.totalSuppDisc, AppTheme.success),
              _row('کسٹمر ٹیکس', bp.totalCustTax, AppTheme.info),
              _divider(),
              _row('خالص منافع', bp.netProfit, bp.netProfit >= 0 ? AppTheme.success : AppTheme.danger, bold: true, large: true),
            ],
          ),
          const SizedBox(height: 14),

          // آج کا منافع
          _profitCard(
            '📅 آج کا تجزیہ',
            [
              _row('آج فروخت', bp.todaySalesTotal, AppTheme.primary),
              _row('آج منافع (مجموعی)', bp.todayGrossProfit, AppTheme.success),
              _row('آج اخراجات', -bp.todayExpTotal, AppTheme.danger),
              _row('آج کسٹمر ڈسکاؤنٹ', -0.0, AppTheme.warning),
              _row('آج سپلائر ڈسکاؤنٹ', 0.0, AppTheme.success),
              _row('آج کسٹمر ٹیکس', 0.0, AppTheme.info),
              _divider(),
              _row('آج خالص منافع',
                bp.todayGrossProfit - bp.todayExpTotal - 0.0 + 0.0 + 0.0,
                AppTheme.success, bold: true),
            ],
          ),
          const SizedBox(height: 14),

          // مجموعی اعداد
          _profitCard(
            '📦 کاروباری خلاصہ',
            [
              _row('کل فروخت (مالیت)', bp.sales.fold(0.0, (s, x) => s + x.total), AppTheme.success),
              _row('کل خریداری (مالیت)', bp.purchases.fold(0.0, (s, x) => s + x.total), AppTheme.warning),
              _row('کسٹمر واجبات', bp.totalReceivables, AppTheme.danger),
              _row('سپلائر واجبات', bp.totalPayables, AppTheme.danger),
              _row('اسٹاک مالیت', bp.totalStockValue, AppTheme.info),
              _row('کیش ان ہینڈ', bp.cashInHand, AppTheme.primary),
            ],
          ),
        ],
      ),
    );
  }

  // ====== ماہانہ ٹیب ======
  Widget _buildMonthlyTab(BusinessProvider bp) {
    final months = [
      'جنوری', 'فروری', 'مارچ', 'اپریل', 'مئی', 'جون',
      'جولائی', 'اگست', 'ستمبر', 'اکتوبر', 'نومبر', 'دسمبر'
    ];

    return Column(
      children: [
        // سال انتخاب
        Container(
          padding: const EdgeInsets.all(10),
          color: AppTheme.success,
          child: Row(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              IconButton(
                icon: const Icon(Icons.arrow_back, color: Colors.white),
                onPressed: () => setState(() => _selectedYear--),
              ),
              Text('$_selectedYear',
                  style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 18)),
              IconButton(
                icon: const Icon(Icons.arrow_forward, color: Colors.white),
                onPressed: () => setState(() => _selectedYear++),
              ),
            ],
          ),
        ),
        Expanded(
          child: ListView.builder(
            padding: const EdgeInsets.all(12),
            itemCount: 12,
            itemBuilder: (ctx, i) {
              final month = i + 1;
              final profit = bp.getMonthlyProfit(month, _selectedYear);

              // اس ماہ کی فروخت اور اخراجات
              double sales = 0, expenses = 0;
              for (final s in bp.sales) {
                final d = DateTime.tryParse(s.date);
                if (d != null && d.month == month && d.year == _selectedYear) sales += s.total;
              }
              for (final e in bp.expenses) {
                final d = DateTime.tryParse(e.date);
                if (d != null && d.month == month && d.year == _selectedYear) expenses += e.amount;
              }

              final isCurrentMonth = month == DateTime.now().month && _selectedYear == DateTime.now().year;

              return Card(
                margin: const EdgeInsets.only(bottom: 8),
                color: isCurrentMonth ? AppTheme.primary.withOpacity(0.05) : null,
                child: Container(
                  decoration: isCurrentMonth
                      ? BoxDecoration(
                          borderRadius: BorderRadius.circular(12),
                          border: Border.all(color: AppTheme.primary.withOpacity(0.4), width: 2),
                        )
                      : null,
                  child: Padding(
                    padding: const EdgeInsets.all(12),
                    child: Row(
                      children: [
                        SizedBox(
                          width: 70,
                          child: Text(months[i],
                              style: TextStyle(
                                  fontWeight: FontWeight.bold,
                                  color: isCurrentMonth ? AppTheme.primary : null)),
                        ),
                        Expanded(
                          child: Column(
                            children: [
                              // بار گراف
                              if (sales > 0 || expenses > 0) ...[
                                _miniBar('فروخت', sales,
                                    bp.sales.fold(0.0, (s, x) => s + x.total),
                                    AppTheme.success),
                                const SizedBox(height: 3),
                                _miniBar('اخراجات', expenses,
                                    bp.totalExpenses.clamp(1, double.infinity),
                                    AppTheme.danger),
                              ],
                            ],
                          ),
                        ),
                        const SizedBox(width: 12),
                        Column(
                          crossAxisAlignment: CrossAxisAlignment.end,
                          children: [
                            Text('Rs. ${sales.toStringAsFixed(0)}',
                                style: const TextStyle(fontSize: 11, color: AppTheme.success)),
                            Text(
                              profit >= 0
                                  ? '+Rs. ${profit.toStringAsFixed(0)}'
                                  : '-Rs. ${profit.abs().toStringAsFixed(0)}',
                              style: TextStyle(
                                  fontWeight: FontWeight.bold,
                                  color: profit >= 0 ? AppTheme.success : AppTheme.danger,
                                  fontSize: 13),
                            ),
                          ],
                        ),
                      ],
                    ),
                  ),
                ),
              );
            },
          ),
        ),
      ],
    );
  }

  Widget _miniBar(String label, double value, double max, Color color) {
    final pct = max > 0 ? (value / max).clamp(0.0, 1.0) : 0.0;
    return Row(
      children: [
        SizedBox(width: 40, child: Text(label, style: const TextStyle(fontSize: 9))),
        Expanded(
          child: ClipRRect(
            borderRadius: BorderRadius.circular(4),
            child: LinearProgressIndicator(
              value: pct,
              backgroundColor: color.withOpacity(0.15),
              valueColor: AlwaysStoppedAnimation<Color>(color),
              minHeight: 8,
            ),
          ),
        ),
      ],
    );
  }

  // ====== آئٹم وار ٹیب ======
  Widget _buildItemWiseTab(BusinessProvider bp) {
    final items = bp.stockItems.map((item) {
      final profit = bp.calcItemProfit(item.id);
      double totalQtySold = 0;
      double totalRevenue = 0;
      for (final sale in bp.sales) {
        for (final si in sale.items.where((i) => i.itemId == item.id)) {
          totalQtySold += si.qty;
          totalRevenue += si.total;
        }
      }
      return _ItemProfitData(
        item: item,
        profit: profit,
        qtySold: totalQtySold,
        revenue: totalRevenue,
      );
    }).toList()
      ..sort((a, b) => b.profit.compareTo(a.profit));

    if (items.isEmpty) {
      return const EmptyState(message: 'کوئی آئٹم نہیں', icon: Icons.inventory_2_outlined);
    }

    return ListView.builder(
      padding: const EdgeInsets.all(12),
      itemCount: items.length,
      itemBuilder: (ctx, i) {
        final d = items[i];
        final maxProfit = items.first.profit.abs().clamp(1.0, double.infinity);

        return Card(
          margin: const EdgeInsets.only(bottom: 10),
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(
                      child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                        Text(d.item.name,
                            style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15)),
                        Text(
                          'فروخت: ${d.qtySold.toStringAsFixed(1)} ${d.item.unit} | آمدنی: Rs. ${d.revenue.toStringAsFixed(0)}',
                          style: TextStyle(fontSize: 11, color: Colors.grey.shade500),
                        ),
                        Text(
                          'اسٹاک: ${d.item.quantity} ${d.item.unit} | WAC: Rs. ${d.item.purchaseRate.toStringAsFixed(2)}',
                          style: TextStyle(fontSize: 11, color: Colors.grey.shade500),
                        ),
                      ]),
                    ),
                    Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                      Text(
                        '${d.profit >= 0 ? '+' : ''}Rs. ${d.profit.toStringAsFixed(0)}',
                        style: TextStyle(
                          fontWeight: FontWeight.w900,
                          color: d.profit >= 0 ? AppTheme.success : AppTheme.danger,
                          fontSize: 16,
                        ),
                      ),
                      StatusBadge(
                        label: d.profit >= 0 ? '✅ فائدہ' : '❌ نقصان',
                        color: d.profit >= 0 ? AppTheme.success : AppTheme.danger,
                      ),
                    ]),
                  ],
                ),
                const SizedBox(height: 8),
                ClipRRect(
                  borderRadius: BorderRadius.circular(4),
                  child: LinearProgressIndicator(
                    value: (d.profit.abs() / maxProfit).clamp(0.0, 1.0),
                    backgroundColor: (d.profit >= 0 ? AppTheme.success : AppTheme.danger).withOpacity(0.15),
                    valueColor: AlwaysStoppedAnimation<Color>(
                        d.profit >= 0 ? AppTheme.success : AppTheme.danger),
                    minHeight: 8,
                  ),
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  Widget _profitCard(String title, List<Widget> rows) {
    return Card(
      child: Column(
        children: [
          Container(
            width: double.infinity,
            padding: const EdgeInsets.all(12),
            decoration: const BoxDecoration(
              color: AppTheme.success,
              borderRadius: BorderRadius.only(topLeft: Radius.circular(12), topRight: Radius.circular(12)),
            ),
            child: Text(title, style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 15)),
          ),
          Padding(padding: const EdgeInsets.all(12), child: Column(children: rows)),
        ],
      ),
    );
  }

  Widget _row(String label, double value, Color color, {bool bold = false, bool large = false}) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 13)),
          Text(
            'Rs. ${value.toStringAsFixed(2)}',
            style: TextStyle(
              color: color,
              fontWeight: bold ? FontWeight.w900 : FontWeight.bold,
              fontSize: large ? 16 : 13,
            ),
          ),
        ],
      ),
    );
  }

  Widget _divider() => const Divider(height: 12);
}

class _ItemProfitData {
  final StockItem item;
  final double profit;
  final double qtySold;
  final double revenue;
  _ItemProfitData({required this.item, required this.profit, required this.qtySold, required this.revenue});
}

// ============================================================
// کیش ان/آؤٹ اسکرین
// ============================================================

class CashInOutScreen extends StatefulWidget {
  const CashInOutScreen({super.key});
  @override
  State<CashInOutScreen> createState() => _CashInOutScreenState();
}

class _CashInOutScreenState extends State<CashInOutScreen>
    with SingleTickerProviderStateMixin {
  late TabController _tabs;

  @override
  void initState() {
    super.initState();
    _tabs = TabController(length: 2, vsync: this);
  }

  @override
  void dispose() {
    _tabs.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final entries = bp.cashInOutList;
      final history = bp.cashLedger;

      final totalIn = entries.where((e) => e.amount > 0).fold(0.0, (s, e) => s + e.amount);
      final totalOut = entries.where((e) => e.amount < 0).fold(0.0, (s, e) => s + e.amount.abs());

      return Scaffold(
        appBar: AppBar(
          title: const Text('💰 کیش ان/آؤٹ'),
          backgroundColor: AppTheme.primary,
          bottom: TabBar(
            controller: _tabs,
            indicatorColor: Colors.white,
            labelColor: Colors.white,
            unselectedLabelColor: Colors.white60,
            tabs: const [
              Tab(text: '💵 انٹریاں'),
              Tab(text: '📜 کیش تاریخچہ'),
            ],
          ),
        ),
        body: Column(
          children: [
            // سمری
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
              color: AppTheme.primary,
              child: Row(
                children: [
                  Expanded(child: _chip('کیش ان ہینڈ', 'Rs. ${bp.cashInHand.toStringAsFixed(0)}',
                      Colors.white, large: true)),
                  Expanded(child: _chip('کل ان', 'Rs. ${totalIn.toStringAsFixed(0)}', Colors.green.shade200)),
                  Expanded(child: _chip('کل آؤٹ', 'Rs. ${totalOut.toStringAsFixed(0)}', Colors.red.shade200)),
                ],
              ),
            ),
            Expanded(
              child: TabBarView(
                controller: _tabs,
                children: [
                  // ====== انٹریاں ======
                  Column(
                    children: [
                      Padding(
                        padding: const EdgeInsets.all(12),
                        child: Row(
                          children: [
                            Expanded(
                              child: ElevatedButton.icon(
                                onPressed: () => _showDialog(context, bp, true),
                                icon: const Icon(Icons.add),
                                label: const Text('💵 کیش ان'),
                                style: ElevatedButton.styleFrom(
                                    backgroundColor: AppTheme.success, foregroundColor: Colors.white),
                              ),
                            ),
                            const SizedBox(width: 12),
                            Expanded(
                              child: ElevatedButton.icon(
                                onPressed: () => _showDialog(context, bp, false),
                                icon: const Icon(Icons.remove),
                                label: const Text('💸 کیش آؤٹ'),
                                style: ElevatedButton.styleFrom(
                                    backgroundColor: AppTheme.danger, foregroundColor: Colors.white),
                              ),
                            ),
                          ],
                        ),
                      ),
                      Expanded(
                        child: entries.isEmpty
                            ? const EmptyState(message: 'کوئی انٹری نہیں', icon: Icons.account_balance_wallet_outlined)
                            : ListView.builder(
                                padding: const EdgeInsets.symmetric(horizontal: 12),
                                itemCount: entries.length,
                                itemBuilder: (ctx, i) {
                                  final e = entries[i];
                                  final isIn = e.amount > 0;
                                  return Card(
                                    margin: const EdgeInsets.only(bottom: 8),
                                    child: ListTile(
                                      leading: CircleAvatar(
                                        backgroundColor: (isIn ? AppTheme.success : AppTheme.danger).withOpacity(0.1),
                                        child: Icon(isIn ? Icons.arrow_downward : Icons.arrow_upward,
                                            color: isIn ? AppTheme.success : AppTheme.danger),
                                      ),
                                      title: Text(e.type, style: const TextStyle(fontWeight: FontWeight.bold)),
                                      subtitle: Text('${e.date}${e.note.isNotEmpty ? " | ${e.note}" : ""}',
                                          style: const TextStyle(fontSize: 11)),
                                      trailing: Row(
                                        mainAxisSize: MainAxisSize.min,
                                        children: [
                                          Text(
                                            '${isIn ? '+' : '-'}Rs. ${e.amount.abs().toStringAsFixed(2)}',
                                            style: TextStyle(
                                                color: isIn ? AppTheme.success : AppTheme.danger,
                                                fontWeight: FontWeight.bold),
                                          ),
                                          IconButton(
                                            icon: const Icon(Icons.delete, size: 16, color: AppTheme.danger),
                                            onPressed: () => _deleteEntry(context, bp, e),
                                          ),
                                        ],
                                      ),
                                    ),
                                  );
                                },
                              ),
                      ),
                    ],
                  ),

                  // ====== کیش تاریخچہ ======
                  history.isEmpty
                      ? const EmptyState(message: 'کوئی تاریخچہ نہیں', icon: Icons.history)
                      : ListView.builder(
                          padding: const EdgeInsets.all(12),
                          itemCount: history.length,
                          itemBuilder: (ctx, i) {
                            final h = history[i];
                            final isPositive = h.amount >= 0;
                            return Container(
                              margin: const EdgeInsets.only(bottom: 4),
                              padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                              decoration: BoxDecoration(
                                color: isPositive
                                    ? AppTheme.success.withOpacity(0.04)
                                    : AppTheme.danger.withOpacity(0.04),
                                borderRadius: BorderRadius.circular(8),
                                border: Border(
                                  right: BorderSide(
                                      color: isPositive ? AppTheme.success : AppTheme.danger, width: 3),
                                ),
                              ),
                              child: Row(
                                children: [
                                  Expanded(
                                    child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                                      Text(h.type, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 13)),
                                      Text('${h.date}${h.note.isNotEmpty ? " | ${h.note}" : ""}',
                                          style: TextStyle(fontSize: 10, color: Colors.grey.shade500)),
                                    ]),
                                  ),
                                  Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                                    Text(
                                      '${isPositive ? '+' : ''}Rs. ${h.amount.toStringAsFixed(2)}',
                                      style: TextStyle(
                                          color: isPositive ? AppTheme.success : AppTheme.danger,
                                          fontWeight: FontWeight.bold,
                                          fontSize: 12),
                                    ),
                                    Text('بیلنس: Rs. ${h.balance.toStringAsFixed(2)}',
                                        style: const TextStyle(fontSize: 10, color: Colors.grey)),
                                  ]),
                                ],
                              ),
            ),
                          },
                        ),
                ],
              ),
            ),
          ],
        ),
      );
    });
  }

  Widget _chip(String label, String value, Color color, {bool large = false}) => Column(
    children: [
      Text(value, style: TextStyle(color: color, fontWeight: FontWeight.bold, fontSize: large ? 16 : 13)),
      Text(label, style: const TextStyle(color: Colors.white60, fontSize: 10)),
    ],
  );

  void _showDialog(BuildContext context, BusinessProvider bp, bool isCashIn) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();

    showModalBottomSheet(
      context: context,
      isScrollControlled: true,
      backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(
          key: key,
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text(isCashIn ? '💵 کیش ان' : '💸 کیش آؤٹ',
                  style: TextStyle(
                      fontSize: 18, fontWeight: FontWeight.bold,
                      color: isCashIn ? AppTheme.success : AppTheme.danger)),
              const SizedBox(height: 16),
              AppFormField(
                label: '💰 رقم',
                controller: amountCtrl,
                required: true,
                keyboardType: const TextInputType.numberWithOptions(decimal: true),
                validator: (v) => (v == null || v.isEmpty || (double.tryParse(v) ?? 0) <= 0) ? 'درست رقم درج کریں' : null,
              ),
              const SizedBox(height: 12),
              DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
              const SizedBox(height: 12),
              AppFormField(label: '📝 نوٹ', controller: noteCtrl, hint: 'وجہ یا تفصیل'),
              const SizedBox(height: 20),
              SizedBox(
                width: double.infinity,
                child: ElevatedButton(
                  onPressed: () async {
                    if (!key.currentState!.validate()) return;
                    final amount = double.parse(amountCtrl.text);
                    if (isCashIn) {
                      await bp.addCashIn(amount, dateCtrl.text, noteCtrl.text);
                    } else {
                      await bp.addCashOut(amount, dateCtrl.text, noteCtrl.text);
                    }
                    if (ctx.mounted) Navigator.pop(ctx);
                    if (mounted) showSnackBar(context,
                        '✅ Rs. ${amount.toStringAsFixed(2)} ${isCashIn ? 'ان' : 'آؤٹ'} محفوظ!');
                  },
                  style: ElevatedButton.styleFrom(
                      backgroundColor: isCashIn ? AppTheme.success : AppTheme.danger,
                      foregroundColor: Colors.white),
                  child: const Text('💾 محفوظ'),
                ),
              ),
              const SizedBox(height: 16),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> _deleteEntry(BuildContext context, BusinessProvider bp, CashInOutEntry e) async {
    final ok = await showConfirmDialog(context, '🗑️ انٹری حذف کریں؟',
        '${e.type} — Rs. ${e.amount.abs().toStringAsFixed(2)}');
    if (ok) {
      await bp.deleteCashEntry(e.id);
      if (mounted) showSnackBar(context, '✅ حذف!');
    }
  }
}


// ============================================================
// پارٹنرز اسکرین
// ============================================================

class PartnersScreen extends StatelessWidget {
  const PartnersScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      final totalShares = bp.partners.fold(0.0, (s, p) => s + p.sharePercent);
      return Scaffold(
        appBar: AppBar(
          title: const Text('🤝 پارٹنرز'),
          backgroundColor: AppTheme.purple,
          actions: [
            IconButton(
              icon: const Icon(Icons.add),
              onPressed: () => _showAddDialog(context, bp),
            ),
          ],
        ),
        body: SingleChildScrollView(
          padding: const EdgeInsets.all(14),
          child: Column(
            children: [
              // نیٹ منافع سمری
              GradientSummaryBox(
                title: '📈 کل خالص منافع',
                amount: 'Rs. ${bp.netProfit.toStringAsFixed(2)}',
                icon: Icons.trending_up,
                colors: const [AppTheme.purple, Color(0xFF9c27b0)],
                subtitle: 'یہ رقم پارٹنرز میں تقسیم ہو گی',
              ),
              const SizedBox(height: 14),

              // شیئر مجموعہ انتباہ
              if (totalShares != 100 && bp.partners.isNotEmpty)
                Container(
                  margin: const EdgeInsets.only(bottom: 12),
                  padding: const EdgeInsets.all(10),
                  decoration: BoxDecoration(
                    color: AppTheme.warning.withOpacity(0.15),
                    borderRadius: BorderRadius.circular(8),
                    border: Border.all(color: AppTheme.warning.withOpacity(0.4)),
                  ),
                  child: Row(
                    children: [
                      const Icon(Icons.warning, color: AppTheme.warning, size: 18),
                      const SizedBox(width: 8),
                      Text(
                        'کل شیئر: $totalShares% (100% ہونا چاہیے)',
                        style: const TextStyle(color: AppTheme.warning, fontWeight: FontWeight.bold),
                      ),
                    ],
                  ),
                ),

              // پارٹنرز کارڈز
              if (bp.partners.isEmpty)
                EmptyState(
                  message: 'کوئی پارٹنر نہیں',
                  icon: Icons.people_outline,
                  buttonLabel: '➕ پارٹنر شامل کریں',
                  onButtonPress: () => _showAddDialog(context, bp),
                )
              else
                ...bp.partners.map((p) => _partnerCard(context, bp, p)),
            ],
          ),
        ),
        floatingActionButton: FloatingActionButton(
          onPressed: () => _showAddDialog(context, bp),
          backgroundColor: AppTheme.purple,
          child: const Icon(Icons.person_add, color: Colors.white),
        ),
      );
    });
  }

  Widget _partnerCard(BuildContext context, BusinessProvider bp, Partner p) {
    final share = bp.getPartnerShare(p.id);
    final balance = share - p.totalWithdrawal;

    return Card(
      margin: const EdgeInsets.only(bottom: 12),
      child: Padding(
        padding: const EdgeInsets.all(14),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            // ہیڈر
            Row(
              children: [
                CircleAvatar(
                  radius: 24,
                  backgroundColor: AppTheme.purple.withOpacity(0.1),
                  child: Text(p.name.isNotEmpty ? p.name[0] : '?',
                      style: const TextStyle(fontSize: 20, color: AppTheme.purple, fontWeight: FontWeight.bold)),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                    Text(p.name, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 17)),
                    Row(children: [
                      const Icon(Icons.pie_chart, size: 14, color: AppTheme.purple),
                      const SizedBox(width: 4),
                      Text('${p.sharePercent}% حصہ',
                          style: const TextStyle(color: AppTheme.purple, fontWeight: FontWeight.bold)),
                    ]),
                  ]),
                ),
                PopupMenuButton<String>(
                  onSelected: (v) {
                    if (v == 'edit') _showEditDialog(context, bp, p);
                    if (v == 'delete') _deletePartner(context, bp, p);
                    if (v == 'withdraw') _showWithdrawDialog(context, bp, p);
                  },
                  itemBuilder: (_) => [
                    const PopupMenuItem(value: 'withdraw', child: Text('💸 نکاسی')),
                    const PopupMenuItem(value: 'edit', child: Text('✏️ ترمیم')),
                    const PopupMenuItem(value: 'delete', child: Text('🗑️ حذف', style: TextStyle(color: AppTheme.danger))),
                  ],
                ),
              ],
            ),
            const SizedBox(height: 12),

            // شیئر تفصیل
            _infoGrid([
              ('منافع میں حصہ', 'Rs. ${share.toStringAsFixed(2)}', AppTheme.success),
              ('کل نکاسی', 'Rs. ${p.totalWithdrawal.toStringAsFixed(2)}', AppTheme.danger),
              ('باقی بیلنس', 'Rs. ${balance.toStringAsFixed(2)}',
                  balance >= 0 ? AppTheme.primary : AppTheme.danger),
            ]),
            const SizedBox(height: 10),

            // نکاسی بٹن
            SizedBox(
              width: double.infinity,
              child: ElevatedButton.icon(
                onPressed: () => _showWithdrawDialog(context, bp, p),
                icon: const Icon(Icons.money, size: 16),
                label: const Text('💸 نکاسی کریں'),
                style: ElevatedButton.styleFrom(
                    backgroundColor: AppTheme.danger, foregroundColor: Colors.white),
              ),
            ),

            // نکاسی تاریخچہ
            if (p.withdrawals.isNotEmpty) ...[
              const SizedBox(height: 8),
              ExpansionTile(
                title: Text('📋 نکاسی تاریخچہ (${p.withdrawals.length})',
                    style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)),
                tilePadding: EdgeInsets.zero,
                children: p.withdrawals.take(5).map((w) => Container(
                  padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 6),
                  decoration: BoxDecoration(
                      border: Border(bottom: BorderSide(color: AppTheme.cardBorder.withOpacity(0.4)))),
                  child: Row(children: [
                    Expanded(child: Text('${w['date']} — ${w['note'] ?? ''}',
                        style: const TextStyle(fontSize: 11))),
                    Text('Rs. ${(w['amount'] as num).toStringAsFixed(2)}',
                        style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: AppTheme.danger)),
                  ]),
                )).toList(),
              ),
            ],
          ],
        ),
      ),
    );
  }

  Widget _infoGrid(List<(String, String, Color)> items) {
    return Row(
      children: items.map((item) => Expanded(
        child: Container(
          margin: const EdgeInsets.symmetric(horizontal: 4),
          padding: const EdgeInsets.all(8),
          decoration: BoxDecoration(
            color: item.$3.withOpacity(0.08),
            borderRadius: BorderRadius.circular(8),
            border: Border.all(color: item.$3.withOpacity(0.2)),
          ),
          child: Column(
            children: [
              Text(item.$1, style: TextStyle(fontSize: 9, color: item.$3, fontWeight: FontWeight.bold), textAlign: TextAlign.center),
              const SizedBox(height: 3),
              Text(item.$2, style: TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: item.$3), textAlign: TextAlign.center),
            ],
          ),
        ),
      )).toList(),
    );
  }

  void _showAddDialog(BuildContext context, BusinessProvider bp) {
    final nameCtrl = TextEditingController();
    final shareCtrl = TextEditingController();
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key, child: Column(mainAxisSize: MainAxisSize.min, children: [
          const Text('🤝 نیا پارٹنر', style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
          const SizedBox(height: 16),
          AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
          const SizedBox(height: 12),
          AppFormField(
            label: '📊 شیئر فیصد (%)',
            controller: shareCtrl, required: true,
            keyboardType: const TextInputType.numberWithOptions(decimal: true),
            validator: (v) {
              if (v == null || v.isEmpty) return 'شیئر فیصد ضروری ہے';
              final n = double.tryParse(v) ?? 0;
              if (n <= 0 || n > 100) return '1 سے 100 کے درمیان درج کریں';
              return null;
            },
          ),
          const SizedBox(height: 20),
          SizedBox(width: double.infinity, child: ElevatedButton(
            onPressed: () async {
              if (!key.currentState!.validate()) return;
              await bp.addPartner(Partner(
                id: 0, name: nameCtrl.text.trim(),
                sharePercent: double.parse(shareCtrl.text),
                totalWithdrawal: 0, withdrawals: [],
              ));
              if (ctx.mounted) Navigator.pop(ctx);
              if (context.mounted) showSnackBar(context, '✅ پارٹنر شامل!');
            },
            style: ElevatedButton.styleFrom(backgroundColor: AppTheme.purple, foregroundColor: Colors.white),
            child: const Text('💾 محفوظ'),
          )),
          const SizedBox(height: 16),
        ])),
      ),
    );
  }

  void _showEditDialog(BuildContext context, BusinessProvider bp, Partner p) {
    final nameCtrl = TextEditingController(text: p.name);
    final shareCtrl = TextEditingController(text: p.sharePercent.toString());
    final key = GlobalKey<FormState>();
    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key, child: Column(mainAxisSize: MainAxisSize.min, children: [
          Text('✏️ ${p.name}', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
          const SizedBox(height: 16),
          AppFormField(label: '📛 نام', controller: nameCtrl, required: true),
          const SizedBox(height: 12),
          AppFormField(label: '📊 شیئر (%)', controller: shareCtrl, required: true,
              keyboardType: const TextInputType.numberWithOptions(decimal: true)),
          const SizedBox(height: 20),
          SizedBox(width: double.infinity, child: ElevatedButton(
            onPressed: () async {
              if (!key.currentState!.validate()) return;
              await bp.updatePartner(p.copyWith(
                name: nameCtrl.text.trim(),
                sharePercent: double.parse(shareCtrl.text),
              ));
              if (ctx.mounted) Navigator.pop(ctx);
              if (context.mounted) showSnackBar(context, '✅ محفوظ!');
            },
            child: const Text('💾 محفوظ'),
          )),
          const SizedBox(height: 16),
        ])),
      ),
    );
  }

  void _showWithdrawDialog(BuildContext context, BusinessProvider bp, Partner p) {
    final amountCtrl = TextEditingController();
    final noteCtrl = TextEditingController();
    final dateCtrl = TextEditingController(text: DateTime.now().toIso8601String().split('T')[0]);
    final key = GlobalKey<FormState>();
    final share = bp.getPartnerShare(p.id);
    final remaining = share - p.totalWithdrawal;

    showModalBottomSheet(
      context: context, isScrollControlled: true, backgroundColor: Colors.white,
      shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
      builder: (ctx) => Padding(
        padding: EdgeInsets.only(bottom: MediaQuery.of(ctx).viewInsets.bottom, left: 20, right: 20, top: 20),
        child: Form(key: key, child: Column(mainAxisSize: MainAxisSize.min, children: [
          Text('💸 ${p.name} — نکاسی', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
          const SizedBox(height: 8),
          Container(
            padding: const EdgeInsets.all(10),
            decoration: BoxDecoration(color: AppTheme.info.withOpacity(0.1), borderRadius: BorderRadius.circular(8)),
            child: Column(children: [
              _infoRow('کل حصہ', 'Rs. ${share.toStringAsFixed(2)}'),
              _infoRow('کل نکاسی', 'Rs. ${p.totalWithdrawal.toStringAsFixed(2)}'),
              _infoRow('دستیاب', 'Rs. ${remaining.toStringAsFixed(2)}'),
            ]),
          ),
          const SizedBox(height: 12),
          AppFormField(
            label: '💰 نکاسی رقم',
            controller: amountCtrl, required: true,
            keyboardType: const TextInputType.numberWithOptions(decimal: true),
            validator: (v) {
              if (v == null || v.isEmpty || (double.tryParse(v) ?? 0) <= 0) return 'درست رقم درج کریں';
              return null;
            },
          ),
          const SizedBox(height: 12),
          DatePickerField(label: '📅 تاریخ', controller: dateCtrl),
          const SizedBox(height: 12),
          AppFormField(label: '📝 نوٹ', controller: noteCtrl),
          const SizedBox(height: 20),
          SizedBox(width: double.infinity, child: ElevatedButton(
            onPressed: () async {
              if (!key.currentState!.validate()) return;
              await bp.partnerWithdrawal(p.id, double.parse(amountCtrl.text), noteCtrl.text, dateCtrl.text);
              if (ctx.mounted) Navigator.pop(ctx);
              if (context.mounted) showSnackBar(context, '✅ نکاسی محفوظ!');
            },
            style: ElevatedButton.styleFrom(backgroundColor: AppTheme.danger, foregroundColor: Colors.white),
            child: const Text('💾 محفوظ'),
          )),
          const SizedBox(height: 16),
        ])),
      ),
    );
  }

  Widget _infoRow(String l, String v) => Padding(
    padding: const EdgeInsets.symmetric(vertical: 3),
    child: Row(mainAxisAlignment: MainAxisAlignment.spaceBetween, children: [
      Text(l, style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold)),
      Text(v, style: const TextStyle(fontSize: 12)),
    ]),
  );

  Future<void> _deletePartner(BuildContext context, BusinessProvider bp, Partner p) async {
    final ok = await showConfirmDialog(context, '🗑️ ${p.name} حذف کریں؟', 'پارٹنر اور تمام نکاسی تاریخچہ حذف ہو جائے گا!');
    if (ok) {
      await bp.deletePartner(p.id);
      if (context.mounted) showSnackBar(context, '✅ حذف!');
    }
  }
}


// ============================================================
// سیٹنگز اسکرین — بیک اپ، ریسٹور، پاسورڈ، ریسیٹ
// ============================================================

class SettingsScreen extends StatelessWidget {
  const SettingsScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Consumer<BusinessProvider>(builder: (context, bp, _) {
      return Scaffold(
        appBar: AppBar(
          title: const Text('⚙️ سیٹنگز'),
          backgroundColor: const Color(0xFF343a40),
        ),
        body: SingleChildScrollView(
          padding: const EdgeInsets.all(14),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // ====== شاپ سیٹنگز ======
              _sectionHeader('🏪 دکان کی سیٹنگز'),
              Card(child: Column(children: [
                _settingTile(
                  icon: Icons.store,
                  color: AppTheme.primary,
                  title: 'دکان کا نام',
                  subtitle: bp.shopName,
                  onTap: () => _changeShopName(context, bp),
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.lock,
                  color: AppTheme.warning,
                  title: 'پاسورڈ تبدیل کریں',
                  subtitle: '****',
                  onTap: () => _changePassword(context, bp),
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.account_balance_wallet,
                  color: AppTheme.success,
                  title: 'ابتدائی کیش بیلنس',
                  subtitle: 'Rs. ${bp.cashInHand.toStringAsFixed(2)}',
                  onTap: () => _setOpeningBalance(context, bp),
                ),
              ])),
              const SizedBox(height: 16),

              // ====== بیک اپ ======
              _sectionHeader('💾 بیک اپ اور ریسٹور'),
              Card(child: Column(children: [
                _settingTile(
                  icon: Icons.backup,
                  color: AppTheme.success,
                  title: '📤 بیک اپ بنائیں',
                  subtitle: 'مکمل ڈیٹا JSON فائل میں محفوظ کریں',
                  onTap: () => _exportBackup(context, bp),
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.restore,
                  color: AppTheme.info,
                  title: '📥 بیک اپ ریسٹور کریں',
                  subtitle: 'JSON فائل سے ڈیٹا واپس لائیں',
                  onTap: () => _importBackup(context, bp),
                ),
              ])),
              const SizedBox(height: 16),

              // ====== ڈیٹا ======
              _sectionHeader('🗄️ ڈیٹا مینجمنٹ'),
              Card(child: Column(children: [
                _settingTile(
                  icon: Icons.cleaning_services,
                  color: AppTheme.warning,
                  title: '🧹 ٹرانزیکشنز صاف کریں',
                  subtitle: 'تمام فروخت، خریداری، اخراجات صاف (اسٹاک/کسٹمر باقی)',
                  onTap: () => _clearTransactions(context, bp),
                  trailing: const Icon(Icons.warning, color: AppTheme.warning, size: 18),
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.delete_forever,
                  color: AppTheme.danger,
                  title: '⚠️ مکمل ریسیٹ',
                  subtitle: 'تمام ڈیٹا ہمیشہ کے لیے حذف',
                  onTap: () => _fullReset(context, bp),
                  trailing: const Icon(Icons.dangerous, color: AppTheme.danger, size: 18),
                ),
              ])),
              const SizedBox(height: 16),

              // ====== معلومات ======
              _sectionHeader('ℹ️ معلومات'),
              Card(child: Column(children: [
                _settingTile(
                  icon: Icons.info_outline,
                  color: AppTheme.info,
                  title: 'ایپ ورژن',
                  subtitle: 'IQBAL TRADERS v1.0.0',
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.people,
                  color: AppTheme.purple,
                  title: 'کسٹمرز',
                  subtitle: '${bp.customers.length} رجسٹرڈ',
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.local_shipping,
                  color: AppTheme.warning,
                  title: 'سپلائرز',
                  subtitle: '${bp.suppliers.length} رجسٹرڈ',
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.inventory_2,
                  color: AppTheme.info,
                  title: 'اسٹاک آئٹمز',
                  subtitle: '${bp.stockItems.length} آئٹمز | مالیت: Rs. ${bp.totalStockValue.toStringAsFixed(0)}',
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.shopping_cart,
                  color: AppTheme.success,
                  title: 'کل فروخت',
                  subtitle: '${bp.sales.length} ٹرانزیکشن',
                ),
                const Divider(height: 1),
                _settingTile(
                  icon: Icons.receipt_long,
                  color: AppTheme.danger,
                  title: 'کل اخراجات',
                  subtitle: 'Rs. ${bp.totalExpenses.toStringAsFixed(2)}',
                ),
              ])),
              const SizedBox(height: 16),

              // ====== لاگ آؤٹ ======
              SizedBox(
                width: double.infinity,
                child: OutlinedButton.icon(
                  onPressed: () => _logout(context),
                  icon: const Icon(Icons.logout),
                  label: const Text('🚪 لاگ آؤٹ'),
                  style: OutlinedButton.styleFrom(
                    foregroundColor: AppTheme.danger,
                    side: const BorderSide(color: AppTheme.danger),
                    padding: const EdgeInsets.symmetric(vertical: 14),
                  ),
                ),
              ),
              const SizedBox(height: 30),
              const Center(
                child: Text(
                  '© IQBAL TRADERS — Built with Flutter',
                  style: TextStyle(color: Colors.grey, fontSize: 12),
                ),
              ),
            ],
          ),
        ),
      );
    });
  }

  Widget _sectionHeader(String title) => Padding(
    padding: const EdgeInsets.only(bottom: 8, top: 4),
    child: Text(title, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15, color: AppTheme.textPrimary)),
  );

  Widget _settingTile({
    required IconData icon,
    required Color color,
    required String title,
    required String subtitle,
    VoidCallback? onTap,
    Widget? trailing,
  }) {
    return ListTile(
      leading: Container(
        padding: const EdgeInsets.all(8),
        decoration: BoxDecoration(color: color.withOpacity(0.1), borderRadius: BorderRadius.circular(8)),
        child: Icon(icon, color: color, size: 20),
      ),
      title: Text(title, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 14)),
      subtitle: Text(subtitle, style: const TextStyle(fontSize: 12)),
      trailing: trailing ?? (onTap != null ? const Icon(Icons.chevron_right, color: Colors.grey) : null),
      onTap: onTap,
    );
  }

  // ====== دکان کا نام ======
  void _changeShopName(BuildContext context, BusinessProvider bp) {
    final ctrl = TextEditingController(text: bp.shopName);
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('🏪 دکان کا نام تبدیل کریں'),
        content: TextFormField(
          controller: ctrl,
          textAlign: TextAlign.right,
          decoration: const InputDecoration(hintText: 'دکان کا نام'),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('منسوخ')),
          ElevatedButton(
            onPressed: () async {
              if (ctrl.text.trim().isNotEmpty) {
                await bp.saveShopName(ctrl.text.trim());
                if (ctx.mounted) Navigator.pop(ctx);
                if (context.mounted) showSnackBar(context, '✅ نام تبدیل!');
              }
            },
            child: const Text('💾 محفوظ'),
          ),
        ],
      ),
    );
  }

  // ====== پاسورڈ ======
  void _changePassword(BuildContext context, BusinessProvider bp) {
    final oldCtrl = TextEditingController();
    final newCtrl = TextEditingController();
    final confirmCtrl = TextEditingController();
    final key = GlobalKey<FormState>();
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('🔐 پاسورڈ تبدیل کریں'),
        content: Form(
          key: key,
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            TextFormField(
              controller: oldCtrl, obscureText: true,
              textAlign: TextAlign.right,
              decoration: const InputDecoration(labelText: '🔑 پرانا پاسورڈ'),
              validator: (v) => v == null || v.isEmpty ? 'پرانا پاسورڈ درج کریں' : null,
            ),
            const SizedBox(height: 10),
            TextFormField(
              controller: newCtrl, obscureText: true,
              textAlign: TextAlign.right,
              decoration: const InputDecoration(labelText: '🆕 نیا پاسورڈ'),
              validator: (v) => v == null || v.length < 4 ? 'کم از کم 4 حروف' : null,
            ),
            const SizedBox(height: 10),
            TextFormField(
              controller: confirmCtrl, obscureText: true,
              textAlign: TextAlign.right,
              decoration: const InputDecoration(labelText: '✅ تصدیق'),
              validator: (v) => v != newCtrl.text ? 'پاسورڈ مماثل نہیں' : null,
            ),
          ]),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('منسوخ')),
          ElevatedButton(
            onPressed: () async {
              if (!key.currentState!.validate()) return;
              final valid = await bp.verifyPassword(oldCtrl.text);
              if (!valid) {
                if (ctx.mounted) showSnackBar(ctx, '❌ پرانا پاسورڈ غلط!', isError: true);
                return;
              }
              await bp.changePassword(newCtrl.text);
              if (ctx.mounted) Navigator.pop(ctx);
              if (context.mounted) showSnackBar(context, '✅ پاسورڈ تبدیل ہو گیا!');
            },
            child: const Text('💾 محفوظ'),
          ),
        ],
      ),
    );
  }

  // ====== ابتدائی بیلنس ======
  void _setOpeningBalance(BuildContext context, BusinessProvider bp) {
    final ctrl = TextEditingController(text: bp.cashInHand.toStringAsFixed(2));
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('💰 کیش ان ہینڈ سیٹ کریں'),
        content: Column(mainAxisSize: MainAxisSize.min, children: [
          const Text('⚠️ یہ براہ راست کیش بیلنس تبدیل کرے گا',
              style: TextStyle(color: AppTheme.warning, fontSize: 12)),
          const SizedBox(height: 10),
          TextFormField(
            controller: ctrl,
            keyboardType: const TextInputType.numberWithOptions(decimal: true),
            textAlign: TextAlign.right,
            decoration: const InputDecoration(labelText: '💰 رقم (Rs.)', border: OutlineInputBorder()),
          ),
        ]),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('منسوخ')),
          ElevatedButton(
            onPressed: () async {
              final amount = double.tryParse(ctrl.text) ?? 0;
              await bp.setOpeningBalance(amount);
              if (ctx.mounted) Navigator.pop(ctx);
              if (context.mounted) showSnackBar(context, '✅ کیش بیلنس: Rs. ${amount.toStringAsFixed(2)}');
            },
            style: ElevatedButton.styleFrom(backgroundColor: AppTheme.success),
            child: const Text('💾 محفوظ'),
          ),
        ],
      ),
    );
  }

  // ====== بیک اپ ======
  Future<void> _exportBackup(BuildContext context, BusinessProvider bp) async {
    try {
      showDialog(
        context: context,
        barrierDismissible: false,
        builder: (_) => const AlertDialog(
          content: Row(children: [
            CircularProgressIndicator(color: AppTheme.primary),
            SizedBox(width: 16),
            Text('بیک اپ بن رہا ہے...'),
          ]),
        ),
      );
      final path = await bp.exportBackup();
      if (!context.mounted) return;
      if (context.mounted) {
        Navigator.pop(context);
        await Share.shareXFiles([XFile(path)], text: 'IQBAL TRADERS بیک اپ');
        showSnackBar(context, '✅ بیک اپ تیار: $path');
      }
    } catch (e) {
      if (context.mounted) {
        Navigator.pop(context);
        showSnackBar(context, '❌ بیک اپ میں خرابی: $e', isError: true);
      }
    }
  }

  // ====== ریسٹور ======
  Future<void> _importBackup(BuildContext context, BusinessProvider bp) async {
    final ok = await showConfirmDialog(
        context, '📥 بیک اپ ریسٹور کریں؟',
        '⚠️ موجودہ تمام ڈیٹا حذف ہو کر بیک اپ سے بحال ہو گا!',
        confirmColor: AppTheme.info);
    if (!ok) return;

    try {
      showDialog(
        context: context,
        barrierDismissible: false,
        builder: (_) => const AlertDialog(
          content: Row(children: [
            CircularProgressIndicator(color: AppTheme.info),
            SizedBox(width: 16),
            Text('ریسٹور ہو رہا ہے...'),
          ]),
        ),
      );
      final success = await bp.importBackup();
      if (!context.mounted) return;
      Navigator.pop(context);
      if (context.mounted) {
        if (success) {
          showSnackBar(context, '✅ ڈیٹا کامیابی سے بحال ہو گیا!');
        } else {
          showSnackBar(context, '❌ فائل منتخب نہیں کی', isError: true);
        }
      }
    } catch (e) {
      if (context.mounted) {
        Navigator.pop(context);
        showSnackBar(context, '❌ ریسٹور میں خرابی: $e', isError: true);
      }
    }
  }

  // ====== ٹرانزیکشنز صاف ======
  Future<void> _clearTransactions(BuildContext context, BusinessProvider bp) async {
    final ok = await showConfirmDialog(
        context, '🧹 ٹرانزیکشنز صاف کریں؟',
        'تمام فروخت، خریداری، اخراجات، اور ادھار کی تاریخچہ حذف ہو گا۔ کسٹمرز اور اسٹاک باقی رہیں گے۔',
        confirmColor: AppTheme.warning);
    if (!ok) return;

    // دوبارہ تصدیق
    final ok2 = await showConfirmDialog(
        context, '⚠️ آخری تصدیق',
        'کیا آپ واقعی یقین سے تمام ٹرانزیکشنز حذف کرنا چاہتے ہیں؟ یہ عمل واپس نہیں ہو سکتا!',
        confirmColor: AppTheme.danger);
    if (!ok2) return;

    try {
      await bp.clearTransactions();
      if (context.mounted) showSnackBar(context, '✅ تمام ٹرانزیکشنز صاف!');
    } catch (e) {
      if (context.mounted) showSnackBar(context, '❌ خرابی: $e', isError: true);
    }
  }

  // ====== مکمل ریسیٹ ======
  Future<void> _fullReset(BuildContext context, BusinessProvider bp) async {
    final passCtrl = TextEditingController();
    final confirmed = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('⚠️ مکمل ریسیٹ', style: TextStyle(color: AppTheme.danger, fontWeight: FontWeight.bold)),
        content: Column(mainAxisSize: MainAxisSize.min, children: [
          const Text('تمام ڈیٹا (کسٹمرز، سپلائرز، اسٹاک، فروخت، خریداری) ہمیشہ کے لیے حذف ہو جائے گا!',
              style: TextStyle(color: AppTheme.danger)),
          const SizedBox(height: 12),
          const Text('تصدیق کے لیے پاسورڈ درج کریں:'),
          const SizedBox(height: 8),
          TextField(
            controller: passCtrl, obscureText: true,
            textAlign: TextAlign.center,
            decoration: const InputDecoration(border: OutlineInputBorder(), hintText: 'پاسورڈ'),
          ),
        ]),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('منسوخ')),
          ElevatedButton(
            onPressed: () => Navigator.pop(ctx, true),
            style: ElevatedButton.styleFrom(backgroundColor: AppTheme.danger),
            child: const Text('ریسیٹ کریں', style: TextStyle(color: Colors.white)),
          ),
        ],
      ),
    ) ?? false;

    if (!confirmed) return;
    final valid = await bp.verifyPassword(passCtrl.text);
    if (!valid) {
      if (context.mounted) showSnackBar(context, '❌ غلط پاسورڈ!', isError: true);
      return;
    }
    await bp.fullReset();
    if (context.mounted) {
      showSnackBar(context, '✅ تمام ڈیٹا حذف!');
      Navigator.pushNamedAndRemoveUntil(context, '/login', (_) => false);
    }
  }

  // ====== لاگ آؤٹ ======
  void _logout(BuildContext context) {
    Navigator.pushNamedAndRemoveUntil(context, '/login', (_) => false);
  }
}


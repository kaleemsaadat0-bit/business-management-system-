import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:provider/provider.dart';
import 'app/app.dart';
import 'app/theme.dart';
import 'providers/business_provider.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  configureGoogleFonts(); // Google Fonts offline-safe

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

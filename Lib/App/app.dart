import 'package:flutter/material.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'package:provider/provider.dart';
import '../providers/business_provider.dart';
import 'theme.dart';
import 'routes.dart';

class IqbalTradersApp extends StatelessWidget {
  const IqbalTradersApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'IQBAL TRADERS',
      debugShowCheckedModeBanner: false,
      theme: AppTheme.theme,
      locale: const Locale('ur', 'PK'),
      supportedLocales: const [
        Locale('ur', 'PK'),
        Locale('en', 'US'),
      ],
      localizationsDelegates: const [
        GlobalMaterialLocalizations.delegate,
        GlobalWidgetsLocalizations.delegate,
        GlobalCupertinoLocalizations.delegate,
      ],
      initialRoute: AppRoutes.login,
      routes: AppRoutes.routes,
      builder: (context, child) {
        return Consumer<BusinessProvider>(
          builder: (context, bp, _) {
            if (bp.isLoading) return _SplashScreen();
            return child ?? const SizedBox.shrink();
          },
        );
      },
    );
  }
}

class _SplashScreen extends StatelessWidget {
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
        child: Center(
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Container(
                padding: const EdgeInsets.all(24),
                decoration: BoxDecoration(
                  color: Colors.white.withOpacity(0.15),
                  shape: BoxShape.circle,
                ),
                child: const Icon(Icons.store_rounded, size: 72, color: Colors.white),
              ),
              const SizedBox(height: 24),
              const Text('IQBAL TRADERS',
                style: TextStyle(color: Colors.white, fontSize: 30,
                    fontWeight: FontWeight.w900, letterSpacing: 2)),
              const SizedBox(height: 8),
              const Text('کاروباری مینجمنٹ سسٹم',
                style: TextStyle(color: Colors.white70, fontSize: 16)),
              const SizedBox(height: 48),
              const CircularProgressIndicator(color: Colors.white, strokeWidth: 2.5),
              const SizedBox(height: 16),
              const Text('لوڈ ہو رہا ہے...',
                style: TextStyle(color: Colors.white60, fontSize: 13)),
            ],
          ),
        ),
      ),
    );
  }
}

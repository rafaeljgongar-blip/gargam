import { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.gargam.stockcontrol',
  appName: 'Gargam Stock',
  webDir: 'dist',
  server: {
    androidScheme: 'https'
  }
};

export default config;

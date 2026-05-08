/** @type {import('tailwindcss').Config} */
export default {
  darkMode: 'class',
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        premium: {
          50: 'hsl(var(--premium-50) / <alpha-value>)',
          100: 'hsl(var(--premium-100) / <alpha-value>)',
          200: 'hsl(var(--premium-200) / <alpha-value>)',
          300: 'hsl(var(--premium-300) / <alpha-value>)',
          400: 'hsl(var(--premium-400) / <alpha-value>)',
          500: 'hsl(var(--premium-500) / <alpha-value>)',
          600: 'hsl(var(--premium-600) / <alpha-value>)',
          700: 'hsl(var(--premium-700) / <alpha-value>)',
          800: 'hsl(var(--premium-800) / <alpha-value>)',
          900: 'hsl(var(--premium-900) / <alpha-value>)',
          950: 'hsl(var(--premium-950) / <alpha-value>)',
        },
      },
      animation: {
        'fade-in': 'fadeIn 0.5s ease-out',
        'slide-up': 'slideUp 0.5s ease-out',
      },
      keyframes: {
        fadeIn: {
          '0%': { opacity: '0' },
          '100%': { opacity: '1' },
        },
        slideUp: {
          '0%': { transform: 'translateY(10px)', opacity: '0' },
          '100%': { transform: 'translateY(0)', opacity: '1' },
        },
      },
    },
  },
  plugins: [],
}

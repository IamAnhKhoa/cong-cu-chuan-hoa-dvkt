import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
    title: 'Chuẩn hóa Danh mục DVKT',
    description: 'Công cụ chuẩn hóa danh mục dịch vụ kỹ thuật theo quy định Bộ Y tế',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
    return (
        <html lang="vi">
            <head>
                <link rel="preconnect" href="https://fonts.googleapis.com" />
                <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
                <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
            </head>
            <body>{children}</body>
        </html>
    );
}

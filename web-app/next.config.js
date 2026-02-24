/** @type {import('next').NextConfig} */
const nextConfig = {
    reactStrictMode: true,
    // Next.js 16+ uses Turbopack by default; xlsx works client-side without extra config
    turbopack: {},
};

module.exports = nextConfig;

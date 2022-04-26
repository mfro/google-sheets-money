import typescript from 'rollup-plugin-typescript2';

module.exports = {
    input: 'src/main.ts',
    output: {
        file: 'out/main.js',
        format: 'cjs',
        intro: 'var exports = {};',
        esModule: false,
    },
    plugins: [
        typescript(),
    ],
};

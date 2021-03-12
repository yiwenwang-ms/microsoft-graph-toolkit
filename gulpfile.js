const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const cleanCSS = require('gulp-clean-css');
const rename = require('gulp-rename');
var license = require('gulp-header-license');

scssFileHeader = (cssLib) => `
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
// MODIFY THE .SCSS FILE INSTEAD

import { css } from '${cssLib}';
/**
 * exports lit-element css
 * @export styles
 */
export const styles = [
  css\``;

scssFileFooter = '`];';

licenseStr = `/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
`;

versionFile = `${licenseStr}
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD

export const PACKAGE_VERSION = '[VERSION]';
`;

function runSassLit() {
  return gulp
    .src('src/**/!(shared)*.scss')
    .pipe(sass())
    .pipe(cleanCSS())
    .pipe(gap.prependText(scssFileHeader('lit-element')))
    .pipe(gap.appendText(scssFileFooter))
    .pipe(rename({ extname: '-css.ts' }))
    .pipe(gulp.dest('src/'));
}

function runSassFast() {
  return gulp
    .src('src/**/!(shared)*.scss')
    .pipe(sass())
    .pipe(cleanCSS())
    .pipe(gap.prependText(scssFileHeader('@microsoft/fast-element')))
    .pipe(gap.appendText(scssFileFooter))
    .pipe(rename({ extname: '-fast-css.ts' }))
    .pipe(gulp.dest('src/'));
}

function runSass() {
  runSassLit();
  return runSassFast();
}

function setLicense() {
  return gulp
    .src(['packages/**/src/**/*.{ts,js,scss}', "!packages/**/generated/**/*"], { base: './' })
    .pipe(license(licenseStr))
    .pipe(gulp.dest('./'));
}

function setVersion() {
  var pkg = require('./package.json');
  var fs = require('fs');
  fs.writeFileSync('./src/utils/version.ts', versionFile.replace('[VERSION]', pkg.version));
}

gulp.task('sass', runSass);
gulp.task('setLicense', setLicense);
gulp.task('setVersion', async () => setVersion());

gulp.task('watchSass', () => {
  runSass();
  return gulp.watch('src/**/*.scss', gulp.series('sass'));
});

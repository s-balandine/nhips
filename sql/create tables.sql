
/* Script DDL de création des tables */
/* .................................................................................... */

CREATE TABLE T_INTERVIEWER_IVW (
    IVW_PRIMARYKEY                   LONG NOT NULL,
    IVW_NAMEFIRST                    VARCHAR(32),
    IVW_NAMELAST                     VARCHAR(32))^

CREATE TABLE T_HOUSEHOLD_HSH (
    HSH_PRIMARY_KEY                  LONG NOT NULL,
    HSH_IDENTIFIER_OF_CLUSTER        VARCHAR(32),
    HSH_IDENTIFIER_OF_HOUSEHOLD      VARCHAR(32),
    HSH_INTERVIEW_SUPERVISOR         LONG,
    HSH_INTERVIEWER                  LONG,
    HSH_SURVEY_TYPIST                LONG,
    HSH_SURVEY_SUPERVISOR            LONG,
    HSH_VISIT_DAY                    INTEGER NOT NULL,
    HSH_VISIT_MONTH                  INTEGER NOT NULL,
    HSH_VISIT_YEAR                   INTEGER NOT NULL,
    HSH_VISIT_RESULT                 SMALLINT NOT NULL)^

CREATE TABLE T_INTERVIEWSUPERVISOR_IVW (
    ISV_PRIMARYKEY                   LONG NOT NULL,
    ISV_NAMEFIRST                    VARCHAR(32),
    ISV_NAMELAST                     VARCHAR(32))^

CREATE TABLE T_SURVEYTYPIST_ISV (
    ISV_PRIMARYKEY                   LONG NOT NULL,
    ISV_NAMEFIRST                    VARCHAR(32),
    ISV_NAMELAST                     VARCHAR(32))^

CREATE TABLE T_SURVEYSUPERVISOR_STP (
    ISV_PRIMARYKEY                   LONG NOT NULL,
    ISV_NAMEFIRST                    VARCHAR(32),
    ISV_NAMELAST                     VARCHAR(32))^



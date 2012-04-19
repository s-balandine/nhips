
/* Script DDL de création pour les contraintes de clé primaire */
/* .................................................................................... */

alter table T_INTERVIEWER_IVW add constraint CKP_T_INTERVIEWER_IVW primary key (IVW_PRIMARYKEY) USING INDEX IKP_T_INTERVIEWER_IVW;

alter table T_HOUSEHOLD_HSH add constraint CKP_T_HOUSEHOLD_HSH primary key (HSH_PRIMARY_KEY) USING INDEX IKP_T_HOUSEHOLD_HSH;

alter table T_INTERVIEWSUPERVISOR_IVW add constraint CKP_T_INTERVIEWSUPERVISOR_IVW primary key (ISV_PRIMARYKEY) USING INDEX IKP_T_INTERVIEWSUPERVISOR_IVW;

alter table T_SURVEYTYPIST_ISV add constraint CKP_T_SURVEYTYPIST_ISV primary key (ISV_PRIMARYKEY) USING INDEX IKP_T_SURVEYTYPIST_ISV;

alter table T_SURVEYSUPERVISOR_STP add constraint CKP_T_SURVEYSUPERVISOR_STP primary key (ISV_PRIMARYKEY) USING INDEX IKP_T_SURVEYSUPERVISOR_STP;


